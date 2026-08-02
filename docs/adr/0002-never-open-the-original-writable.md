# ADR-0002: Never open the original layout writable

**Status:** Accepted (v1.0.0) · **Source:** SOLUTION-DESIGN §6.1, §6.2

## Description

Every mutating tool works on a **staged copy** in the target's own directory, validates the result
(`OpenXmlValidator` diffed against pre-edit state, plus the table-grid and plain-text-nesting guards
for the corruption classes the validator accepts silently), and only then commits with an **atomic
same-volume rename**. A rejected edit (`edit_would_corrupt`, `edit_would_break_table`) leaves the
original byte-identical; a crash mid-edit cannot leave a torn file.

## Why

The file being edited is usually a developer's working copy in a real AL repository — often the only
copy. A Word `.docx` is a zip: a partial write is not "a file with a bad edit", it is an unopenable
container. The asymmetry is total: staging costs one file copy per edit; a torn layout costs a
developer their work plus the confidence to ever point the tool at a real repo again.

## Consequences

- Mutation cost includes a copy of the package per edit — accepted; layouts are small (tens of KB to
  a few MB).
- The staged copy lives in the target's directory *by design*: an atomic rename is only atomic on
  the same volume. Moving staging to a temp dir on another volume reintroduces the torn-file window.
- PRs that open the original with write access "to save a copy" will be declined regardless of
  measured speedup.
