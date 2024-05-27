from oletools.olevba import VBA_Parser
import json
from oletools.mraptor import MacroRaptor

if __name__ == '__main__':
    macro_infos = {}
    vba_parser = VBA_Parser("../assets/VeryBad.docm")
    macro_infos["Macros Keywords"] = {keyword: f"{note}, {description}" for note, keyword, description
                                      in vba_parser.analyze_macros()}
    macro_infos["Macros"] = {path: f"{vba_name}, {code}" for x, path, vba_name, code in
                             vba_parser.extract_all_macros()}
    macro_infos["Form Strings"] = {"": f for f in vba_parser.extract_form_strings()}  # extracts strings from form

    with open("temp.json", "w", encoding="utf-8") as output:
        json.dump(macro_infos, output, indent=4)
