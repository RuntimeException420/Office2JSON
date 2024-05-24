import shutil
import zipfile
import os
import json


def create_json(folder_path):
    data = {}

    for root, d_names, f_names in os.walk(folder_path):  # walks recursively through the file system
        path = root.replace("..\\assets\\temp_extraction\\", "")
        parts = path.split("\\")
        curr_dict = data
        for i, part in enumerate(parts):
            if part == "":
                part = ".\\"
            if i == len(parts):
                curr_dict[part] = ""
            else:
                if part not in curr_dict:
                    curr_dict[part] = {k: read_file_content(root, k) for k in f_names}
                curr_dict = curr_dict[part]
    return data


def read_file_content(path, file):
    file = os.path.join(path, file)
    if file.endswith(".xml") or file.endswith(".rels"):
        with open(file, encoding="utf8") as f:
            return f.read()
    return ""


if __name__ == '__main__':
    try:
        # Duplicate Office-Document and convert it into a zip-file
        shutil.copy("../assets/VeryBad.docm", "../assets/VeryBad.zip")

        zip_copy = zipfile.ZipFile("../assets/VeryBad.zip", "r")
        zip_copy.extractall("../assets/temp_extraction")
        zip_copy.close()

        os.remove("../assets/VeryBad.zip")
        json_dict = create_json("..\\assets\\temp_extraction\\")

        with open("extracted_information.json", "w", encoding="utf-8") as output:
            json.dump(json_dict, output, indent=4)

        shutil.rmtree("../assets/temp_extraction")
    except FileNotFoundError:
        print("File was not found.")
    except shutil.SameFileError:
        print("Error occurred.")
