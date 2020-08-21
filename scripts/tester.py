reg = r"\s+Mo[\s-]*"
reg_2 = r"-+[awt\.0-9%\s]*Cu[\s-]*"
reg_3 = r"-+[awt\.0-9%\s]*Mg[\s-]*"
import re

#trying to ignore other words and look for Mo-Cu or Mo-4%w.t.Cu or Mo-4Cu etc.
if re.search(reg, "hardening of meMo- and Cu as well as Mg---"):
    print("pass")
