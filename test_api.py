import json
from xlsm_safe.spec import apply, extract

if __name__ == "__main__":
    template_path = "data/INSP_240G_PC-CRANE_A_J_VER2300.xlsm"
    spec_path = "examples/pc-crane-spec.json"
    with open(spec_path, "r") as f:
        spec = json.load(f)
    apply(template_path, spec, output="out.xlsm")