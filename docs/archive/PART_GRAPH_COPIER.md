# Part Graph Copier Change Proposal

## Problem

`pptforge` currently copies slide XML byte-for-byte and migrates a whitelist of
known relationships: media, SmartArt diagrams, layouts, notes, and tags. This is
safe, but it means any slide relationship outside that whitelist can be left
pointing at a part that was never copied from a non-base source. Examples include
charts, embedded workbooks, OLE packages, comments, controls, and other Office
extension parts.

OOXML packages are relationship graphs. A slide's complete object payload is not
always contained in `ppt/slides/slideN.xml`; it is usually the closure reachable
from `ppt/slides/_rels/slideN.xml.rels`. A robust slide copy should therefore
copy the reachable internal part graph, then apply targeted rules only where
PowerPoint has global registries or id constraints.

## Desired Direction

The merger should move from a pure relationship whitelist to this shape:

1. Keep slide XML as byte-for-byte pass-through.
2. Preserve specialized strategies for known global or fragile parts:
   - slide ids in `presentation.xml`
   - slide layout / master / theme migration
   - master/layout global id allocation
   - notes slide back-pointers
   - media de-duplication
   - SmartArt diagram grouping
3. For every other internal relationship, recursively copy the target part and
   its own `.rels` tree.
4. Allocate a new output path when the preferred path is already occupied by a
   different source part.
5. Rewrite relationship `Target` values when a copied child part receives a new
   path.
6. Carry over the source content type for every copied part.
7. Leave external relationships unchanged.

## Implementation Notes

This branch introduces `PartGraphCopier` as the default fallback for unknown
internal relationships.

- The copier is path-oriented and never parses or rewrites slide XML.
- It parses only relationship files and `[Content_Types].xml`.
- It caches `(source package, source part path) -> output part path` so shared
  parts stay shared.
- It reserves output paths and suffixes collisions with `_2`, `_3`, and so on.
- It recurses through each copied part's relationship file.
- It delegates known child media and diagram relationships back to
  `MediaManager` and `DiagramManager`, so existing de-duplication and SmartArt
  behavior are preserved.
- It exposes additional content type defaults and overrides for the merger to
  merge into `[Content_Types].xml`.

This is not intended to replace all specialized rules. It is the base layer that
makes "new object type appears in slide relationships" copyable by default,
while special cases remain responsible for package-level invariants.
