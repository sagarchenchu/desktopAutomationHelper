# Step registry (Phase 4D)

Approved, versioned mappings from BDD step text to Phase 2 plan-step templates.

- Schema: `../schemas/step-registry.schema.json`
- Default registry: `registry.json`
- Patterns are either `exact` (with `{param}` placeholders) or bounded `regex`
  (culture-invariant, timeout-bounded, compiled once).
- `risk: prohibited` mappings always fail closed at compile time.
- Object and knowledge references must resolve to **approved** artifacts only.
