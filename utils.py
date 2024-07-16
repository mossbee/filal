import re
import docx
import json

def docx_indexing(file_name):

    doc = docx.Document('data/' + file_name)
    document_data = []

    for para in doc.paragraphs:
        document_data.append({
            "index" : "0.0.0",
            "text" : para.text
        })

    document_data_len = len(document_data)
    data_for_print = []

    for i in range(1, document_data_len):
        previous_element_index_list = document_data[i-1]["index"].split(".")
        previous_element_section_number = int(previous_element_index_list[0])
        previous_element_subsection_number = int(previous_element_index_list[1])
        element = document_data[i]
        if element["text"].startswith("Chương"):
            # remove that element from the list
            continue
        # elif element["text"] starts with "Mục " and a number
        elif re.match(r"^Mục \d+", element["text"]):
            # remove that element from the list
            continue
        elif element["text"].startswith("Điều"):
            # section number is the first word right after "Điều", remove the dot at the end
            section_number = element["text"].split()[1][:-1]
            # update the index
            index_list = element["index"].split(".")
            index_list[0] = section_number
            element["index"] = ".".join(index_list)
            data_for_print.append({
                "index" : element["index"],
                "text" : element["text"]
            })
        elif re.match(r"^\d+\.", element["text"]):
            # subsection number is the first word, remove the dot at the end
            subsection_number = element["text"].split(".")[0]
            # update the index
            index_list = element["index"].split(".")
            index_list[1] = subsection_number
            index_list[0] = str(previous_element_section_number)
            element["index"] = ".".join(index_list)
            data_for_print.append({
                "index" : element["index"],
                "text" : element["text"]
            })
        # else if element["text"] starts with a character and a ) then it is a point, update the index
        elif re.match(r"^[a-zđê]\)", element["text"], re.IGNORECASE):
            # point number is the first character, remove the bracket at the end
            point_number = element["text"].split(")")[0]
            # update the index
            index_list = element["index"].split(".")
            index_list[2] = point_number
            index_list[1] = str(previous_element_subsection_number)
            index_list[0] = str(previous_element_section_number)
            element["index"] = ".".join(index_list)
            data_for_print.append({
                "index" : element["index"],
                "text" : element["text"]
            })
        else:
            index_list = element["index"].split(".")
            index_list[0] = str(previous_element_section_number)
            element["index"] = ".".join(index_list)
            data_for_print.append({
                "index" : element["index"],
                "text" : element["text"]
            })
    # save document_data to json file
    # save document_data to json file inside json folder with the same name as the docx file + "_index.json"
    with open('json/' + file_name.split(".")[0] + '_index.json', 'w', encoding="utf-8") as json_file:
        json.dump({
            "document" : data_for_print
        }, json_file, indent=4, ensure_ascii=False)

    # with open('86_2015_ND-CP_292146_only_dieu_index.json', 'w', encoding="utf-8") as json_file:
    #     json.dump({
    #         "document" : data_for_print
    #     }, json_file, indent=4, ensure_ascii=False)