## What & why

<!-- One functionality per PR. What changes for a user of the tools? -->

## Checklist

- [ ] `dotnet test` is green (full suite).
- [ ] New behavior is covered by tests asserting the resulting OOXML (validator checks paired with a
      valid sibling that must not trip them).
- [ ] Snapshot re-baselines (if any): every diff reviewed individually, none blanket-accepted.
- [ ] Corpus additions (if any): scrubbed with `python tools/scrub_corpus_metadata.py` and recorded
      in `tests/corpus/PROVENANCE.md`; nothing confidential.
- [ ] Conforms to the ADRs in `docs/adr/` (envelope contract, refuse-rather-than-guess,
      staged-copy writes, observed-OOXML-only) — or the PR updates the ADR/design doc in the same
      change and says why.
- [ ] `CHANGELOG.md` updated for user-visible changes.
