"""Interactive Windows acceptance. This test does not install a driver or create a session."""
import argparse
import json
import subprocess
import sys
import time
import traceback
import uuid
from pathlib import Path
from xml.etree import ElementTree as ET


def eventually(predicate, message, seconds=15):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if predicate():
            return
        time.sleep(0.2)
    raise AssertionError(message)


def values(element):
    result = set()
    for child in [element, *element.descendants()]:
        text = child.window_text()
        if text:
            result.add(text.strip())
        try:
            value = child.iface_value.CurrentValue
        except Exception:
            continue  # ValuePattern is optional; Text/name remains available.
        if value:
            result.add(str(value).strip())
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--fixture', required=True, type=Path)
    parser.add_argument('--display', help=r'Active Windows display name, for example \\.\DISPLAY2')
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--expect-wrong-row', action='store_true', help='Negative control: must fail after Add.')
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    report = {'schemaVersion': 1, 'runId': uuid.uuid4().hex, 'outcome': 'infrastructure-failed',
              'environmentLabel': 'operator-must-record-viewer-and-session-state', 'steps': [],
              'negativeControl': args.expect_wrong_row}
    started = time.monotonic()
    process = None
    window = None
    exit_code = 2
    step_name = 'startup'
    try:
        if sys.platform != 'win32':
            raise RuntimeError('This acceptance test requires interactive Windows.')
        from pywinauto import Application
        from PIL import ImageChops
        import importlib.metadata
        report['dependencies'] = {name: importlib.metadata.version(name) for name in ('pywinauto', 'Pillow')}
        fixture = args.fixture.resolve(strict=True)
        command = [str(fixture)]
        if args.display:
            command.extend(['--display', args.display])
        process = subprocess.Popen(command, cwd=str(fixture.parent))
        report['fixturePid'] = process.pid
        app = Application(backend='uia').connect(process=process.pid, timeout=20)
        window = app.window(title_re=r'DesktopRuntime Fixture - .*')
        window.wait('visible enabled ready', timeout=20)
        window.set_focus()

        def control(identifier):
            return window.child_window(auto_id=identifier).wrapper_object()

        def step(name, action):
            nonlocal step_name
            step_name = name
            at = time.monotonic()
            action()
            report['steps'].append({'name': name, 'outcome': 'passed', 'seconds': round(time.monotonic() - at, 3)})

        def type_text(identifier, text):
            target = control(identifier)
            target.click_input()
            target.type_keys('^a{BACKSPACE}')
            target.type_keys(text, with_spaces=True, pause=0.03)
            eventually(lambda: target.iface_value.CurrentValue == text, f'{identifier} did not receive typed input')

        grid = control('RecordsGrid')

        def matching_rows(name, age, role):
            expected = {name, str(age), role}
            return [row for row in grid.descendants(control_type='DataItem') if expected <= values(row)]

        def assert_record(name, age, role):
            eventually(lambda: len(matching_rows(name, age, role)) == 1, f'Expected exactly one grid row: {name}, {age}, {role}')

        def disabled_field():
            if control('DisabledInput').is_enabled():
                raise AssertionError('DisabledInput unexpectedly enabled')

        def empty_name_validation():
            control('AddRecord').click_input()
            eventually(lambda: 'Name must contain' in control('ResultText').window_text(), 'Missing required-name validation')
            if grid.descendants(control_type='DataItem'):
                raise AssertionError('Invalid record was added')

        def fresh_frames():
            marker = control('FrameMarker')
            before_text = marker.window_text()
            before = marker.capture_as_image()
            before.save(args.output / 'marker-before.png')
            eventually(lambda: marker.window_text() != before_text, 'UI marker is not advancing')
            # Compare the marker crop, so a cursor blinking elsewhere cannot satisfy freshness.
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline:
                after = marker.capture_as_image()
                if before.size == after.size and ImageChops.difference(before.convert('RGB'), after.convert('RGB')).getbbox():
                    after.save(args.output / 'marker-after.png')
                    window.capture_as_image().save(args.output / 'fixture.png')
                    return
                time.sleep(0.2)
            raise AssertionError('Capture did not show a changing frame marker')

        step('disabled-control', disabled_field)
        step('required-name-validation', empty_name_validation)
        record_name = 'vdd_' + report['runId'][:8]
        step('real-keyboard-input', lambda: type_text('NameInput', record_name))

        def invalid_age():
            type_text('AgeInput', '0')
            control('AddRecord').click_input()
            eventually(lambda: 'Age must be an integer' in control('ResultText').window_text(), 'Missing numeric range validation')
            if grid.descendants(control_type='DataItem'):
                raise AssertionError('Out-of-range record was added')
            type_text('AgeInput', '30')

        step('numeric-validation', invalid_age)
        step('dropdown', lambda: control('RoleInput').select('Editor'))
        step('add', lambda: control('AddRecord').click_input())
        expected = record_name + '_wrong' if args.expect_wrong_row else record_name
        step('verify-added-row', lambda: assert_record(expected, 30, 'Editor'))
        step('fresh-screenshots', fresh_frames)

        def update():
            matching_rows(record_name, 30, 'Editor')[0].click_input()
            type_text('NameInput', record_name + '_updated')
            type_text('AgeInput', '31')
            control('UpdateRecord').click_input()
            assert_record(record_name + '_updated', 31, 'Editor')
            if record_name in values(grid):
                raise AssertionError('Old row still exists after update')

        step('update-and-verify-grid', update)

        def delete():
            matching_rows(record_name + '_updated', 31, 'Editor')[0].click_input()
            control('DeleteRecord').click_input()
            eventually(lambda: record_name + '_updated' not in values(grid), 'Deleted row remains in grid')
            if grid.descendants(control_type='DataItem'):
                raise AssertionError('Grid is not empty after deletion')

        step('delete-and-verify-absence', delete)
        report['outcome'] = 'passed'
        exit_code = 0
    except Exception as error:
        exit_code = 1 if isinstance(error, AssertionError) else 2
        report['outcome'] = 'assertion-failed' if exit_code == 1 else 'infrastructure-failed'
        report['error'] = str(error)
        report['failedStep'] = step_name
        report['traceback'] = traceback.format_exc()
    finally:
        # This fixture uses only in-memory test records. Never target an enterprise app PID here.
        if process is not None:
            try:
                if window is not None:
                    window.close()
                process.wait(timeout=5)
            except Exception:
                if process.poll() is None:
                    process.terminate()
                    try:
                        process.wait(timeout=5)
                    except Exception as error:
                        report['cleanupError'] = str(error)
                        if exit_code == 0:
                            report['outcome'] = 'cleanup-failed'
                            exit_code = 3
        report['seconds'] = round(time.monotonic() - started, 3)
        (args.output / 'acceptance.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
        suite = ET.Element('testsuite', name='DesktopRuntimeFixture', tests='1', failures=str(int(exit_code == 1)),
                           errors=str(int(exit_code not in (0, 1))), time=str(report['seconds']))
        case = ET.SubElement(suite, 'testcase', classname='DesktopRuntime.VddAcceptance', name='fixture-crud', time=str(report['seconds']))
        if exit_code:
            ET.SubElement(case, 'failure' if exit_code == 1 else 'error', message=report.get('error', report['outcome'])).text = report.get('traceback', '')
        ET.ElementTree(suite).write(args.output / 'junit.xml', encoding='utf-8', xml_declaration=True)
        print(json.dumps({'outcome': report['outcome'], 'report': str(args.output / 'acceptance.json')}))
    return exit_code


if __name__ == '__main__':
    sys.exit(main())
