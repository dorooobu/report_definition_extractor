import argparse
import os.path
import pandas as pd
import xml.dom.minidom

def extract_from_excel(excel_file):
    report_def = pd.read_excel(excel_file)
    report_def = report_def.iloc[10:, [0, 2, 7, 6]]
    report_def.columns = ['field', 'character_count', 'data_type', 'format']
    report_def.to_csv('report_def_excel.csv', index=False)

def extract_from_xml(xml_file):
    doc = xml.dom.minidom.parse(xml_file)
    
    fields = doc.getElementsByTagName("Field")
    field_col = []
    char_count_col = []
    data_type_col = []
    format_col = []
    for field in fields:
        field_col.append(field.getAttribute("name"))
        char_count_col.append(field.getAttribute("charCount"))
        data_type_col.append("文字型" if field.getAttribute("type") == "0" else "数値型")
        format_col.append(field.getAttribute("strEditFormula"))

    report_def = pd.DataFrame(
        {
            "field": field_col,
            "character_count": char_count_col,
            "data_type": data_type_col,
            "format": format_col
        }
    )
    report_def.to_csv('report_def_xml.csv', index=False)

def main():
    description = """
    Extract report definition information from excel file and xml file.
    """.strip()

    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('excel', type=str, help='Report Definition file')
    parser.add_argument('xml', type=str, help='XML report')

    args = parser.parse_args()
    assert os.path.isfile(args.excel), "file {} must exist".format(args.excel)
    assert os.path.isfile(args.xml), "file {} must exist".format(args.xml)

    extract_from_excel(args.excel)
    extract_from_xml(args.xml)

if __name__ == '__main__':
    main()