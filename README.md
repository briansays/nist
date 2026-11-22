# NIST SP 800-53 Rev 5 SCTM builder

Builds `NIST_SP800-53r5_SCTM.xlsx`, a macro-free SCTM, from NIST's official OSCAL data.

## Rebuild

```bash
cd sctm
./fetch_oscal.sh oscal                       # download the latest catalog + 800-53B baselines
python3 -m venv venv && ./venv/bin/pip install openpyxl   # first time only
./venv/bin/python build_sctm.py oscal ../NIST_SP800-53r5_SCTM.xlsx
```

The script prints the number of controls it wrote. The Instructions tab of the
workbook shows the catalog version and baseline counts it was built from.

**Rebuilding makes a new, blank workbook.** It does not keep statuses, overlays or
implementation text from an existing SCTM. Write the output to a new file name
and copy your entries across.

Source data in XLSX format can be found here: 

- https://csrc.nist.rip/publications/detail/sp/800-53/rev-5/final
- https://csrc.nist.rip/publications/detail/sp/800-53b/final
