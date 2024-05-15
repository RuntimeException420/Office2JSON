import shutil
import zipfile
import os
import json


def create_json(folder_path):
    data = {}
    for root, d_names, f_names in os.walk(folder_path):
        for f in f_names:
            print(os.path.join(root, f))
        #for f in f_names:
         #   fname = os.path.join(root, f)
          #  if fname.endswith(".xml") or fname.endswith(".rels"):
           #     with open(fname, encoding="utf8") as temp:
            #        data[fname] = temp.read()

    #print(json.dumps(data, indent=2))


if __name__ == '__main__':
    try:
        """
        # Duplicate Office-Document and convert it into a zip-file
        shutil.copy("../assets/VeryBad.docm", "../assets/VeryBad.zip")
        
        # Decompress archive
        zip_copy = zipfile.ZipFile("../assets/VeryBad.zip", "r")
        zip_copy.extractall("../assets/temp_extraction")
        zip_copy.close()
        
        os.remove("../assets/VeryBad.zip")
        """
        #create_json("..\\assets\\temp_extraction\\")
        with open("../assets/temp_extraction/word/vbaProject.bin/PROJECT", "rb") as f:
            print(f.read())
        print("done")
    #except FileNotFoundError:
    #    print("File was not found.")
    except shutil.SameFileError:
        print("Error occurred.")
