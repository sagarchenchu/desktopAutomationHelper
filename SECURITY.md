# Security Policy

## Supported versions

Security fixes are provided for the latest release and the current `main`
branch. Older releases may not receive patches.

## Reporting a vulnerability

Do not open a public issue for a suspected vulnerability.

Use GitHub's **Report a vulnerability** option on the repository Security page
when private vulnerability reporting is available. Otherwise, contact the
repository owner privately through the contact method on the owner's GitHub
profile.

Include, where practical:

- the affected version or commit;
- the affected endpoint, component, or workflow;
- reproduction steps or a minimal proof of concept;
- the expected impact and required preconditions; and
- suggested mitigations, if known.

Never include real bearer tokens, credentials, customer data, or production
automation artifacts. Use synthetic examples and redact secrets.

You should receive an acknowledgment within seven days. Please allow time to
investigate and prepare a coordinated fix before public disclosure.

## Security boundary

The driver exposes local HTTP endpoints and supports shared Citrix, RDS, and
VDI hosts. Review the documented
[shared-host discovery security boundary](docs/shared-host-driver-discovery-security.md)
before deploying it in a shared environment.
