import shutil
import zipfile
import os
import argparse
import subprocess
import json
import time
import math


def create_json(folder_path):
    data = {}

    for root, d_names, f_names in os.walk(folder_path):  # walks recursively through the file system
        path = root.replace(folder_path, "")
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


def read_file_content(path, file_name):
    file_path = os.path.join(path, file_name)
    if file_path.endswith(".xml") or file_path.endswith(".rels"):
        with open(file_path, encoding="utf8") as f:
            return f.read().replace('"', "'")  # replace quotes so the json.dump does not escape them
    elif file_path.endswith(".bin"):
        try:
            output = subprocess.check_output(["olevba", "--json", file_path]).decode("utf-8")
            clean_output = output[output.find('container') - 1:].replace("\r\n", "")
            clean_output = '{"' + clean_output[1:]
            clean_output = clean_output[:clean_output.rindex("}") + 1]
            clean_output = json.loads(clean_output)
            return clean_output
        except:
            return ""
    return ""


def extract(file_path, output_path: str = None):
    if output_path is None:
        output_path = file_path

    rel_path, file = os.path.split(file_path.file)
    rel_path += "\\"
    temp_name = os.path.splitext(file)[0] + ".zip"

    # Duplicate Office-Document and convert it into a zip-file
    shutil.copy(file_path.file, rel_path + temp_name)

    zip_copy = zipfile.ZipFile(rel_path + temp_name, "r")
    zip_copy.extractall(rel_path + "temp_extraction")
    zip_copy.close()

    os.remove(rel_path + temp_name)
    json_dict = create_json(rel_path + "temp_extraction")
    with open(os.path.join(rel_path, "extracted_" + file + ".json"), "w", encoding="utf-8") as res_file:
        json.dump(json_dict, res_file, indent=4)


if __name__ == '__main__':
    start_time = time.time()
    rel_path = ".\\"
    try:
        args_parser = argparse.ArgumentParser("Office2JSON.py")
        args_parser.add_argument("file", help="Provide the file path to a Microsoft Office document to be analysed.")
        args = args_parser.parse_args()

        rel_path, _ = os.path.split(args.file)
        extract(args)

    except FileNotFoundError as e:
        print("File was not found.\n" + e)
    except shutil.SameFileError as e:
        print("Error occurred.\n" + e)
    except zipfile.BadZipfile as e:
        print("Can't open file as archive.\n")
    finally:
        # clean-up
        file = os.path.join(rel_path, "temp_extraction")
        if os.path.exists(file):
            shutil.rmtree(file)

        print(40 * "_")
        print(f"Extraction time: \t{round(time.time() - start_time, 3)} seconds")
        print(f"Extraction path: \t{rel_path + '/'}")
        print(40 * "_")

