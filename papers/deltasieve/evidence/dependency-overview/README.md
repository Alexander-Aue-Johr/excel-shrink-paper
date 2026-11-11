# Tested dependency example for Figure 2

The figure reconstructs the `dynamic-indirect` control schematically. The native
receipt is pinned by the paper's `indirectcontrols` reference. `case.json` retains
the selected/reference masks, image hashes, outcomes and unsupported-rule decision.

Both input packages were regenerated with the existing `dependency_controls.py`
fixture generator and matched the original experiment's SHA-256 identities.
Their XML confirms X1 changing from 1 to 9, unchanged Z1 containing `X1`, and
the unchanged `INDIRECT($Z$1)>5` rule over A1:A3. The drawing compresses columns
D:W and empty rows 4:29. It depicts the recorded native effect and conservative
fallback. The matched masks apply within the declared A1:AA30 reference.
