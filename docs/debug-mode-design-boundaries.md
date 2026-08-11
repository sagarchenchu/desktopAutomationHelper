# Debug Mode design boundaries

Debug Mode (including future Ctrl+D) is **not implemented** in Phase 3.5.

## Do not grow the overlay or UiService for Debug Mode

Future Ctrl+D work must **not** be placed directly into the large
`RecordingOverlayWindow` or `UiService` classes.

The overlay should remain a **UI/menu coordinator only**.

## Preferred service split

When Debug Mode is implemented, separate it into focused services such as:

- debug-session coordinator
- UIA tree capture service
- locator-candidate generator
- non-interactive highlight service
- candidate verification service
- artifact writer

## This phase

Phase 3.5 records the boundary only. It does **not** perform the large refactor
or ship Debug Mode behavior.
