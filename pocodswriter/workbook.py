import os
import zipfile
import xml.etree.cElementTree as ET

from .format import Format


class Workbook:
    def __init__(self, filename, options):
        self.filename = filename
        self.options = options
        self.worksheets = []
        self.formats = {}

    def add_worksheet(self, name, worksheet_constructor):
        worksheet = worksheet_constructor()
        self.worksheets.append({"name": name, "worksheet": worksheet})
        return worksheet

    def add_format(self, props):
        frozen_props = tuple(sorted(props.items()))
        if frozen_props in self.formats:
            return self.formats[frozen_props]

        format = Format(props, 'format_' + str(len(self.formats)))
        self.formats[frozen_props] = format
        return format

    def _generate_hyperlink_et(self, p_element_et, cell_data):
        url = cell_data["url"]
        if url.startswith('internal:'):
            url = url.replace('internal:', "#").replace('!', '.')
        a_element = ET.SubElement(
            p_element_et,
            "text:a",
            {
                "xlink:href": url,
                "xlink:type": "simple"
            }
        )
        a_element.text = str(cell_data["data"])

    def _generate_table_cell_et(self, table_row_et, cell_data):
        props = {"office:value-type": "string", "calcext:value-type": "string"}
        _format = cell_data["format"]
        if _format is not None:
            props["table:style-name"] = _format.name
        table_cell = ET.SubElement(
            table_row_et,
            "table:table-cell",
            props
        )
        if "url" in cell_data:
            p_element = ET.SubElement(table_cell, "text:p")
            self._generate_hyperlink_et(p_element, cell_data)
        else:
            for paragraph in str(cell_data["data"]).split('\n'):
                p_element = ET.SubElement(table_cell, "text:p")
                p_element.text = paragraph

    def _generate_table_row_et(self, table_et, cell_contents):
        table_row = ET.SubElement(
            table_et,
            "table:table-row",
            {"table:style-name": "ro2"}
        )
        for cell_data in cell_contents:
            self._generate_table_cell_et(table_row, cell_data)

    def _generate_sheet_et(self, worksheet_name, worksheet_index, rows, number_of_columns):
        sheet_et = ET.Element(
            "table:table",
            {"table:name": worksheet_name, "table:style-name": "ta1"}
        )
        for column_index in range(number_of_columns):
            ET.SubElement(
                sheet_et,
                "table:table-column",
                {
                    "table:style-name": self._column_style_name(worksheet_index, column_index),
                    "table:default-cell-style-name": "Default"
                }
            )

        for row in rows:
            self._generate_table_row_et(sheet_et, row)
        return sheet_et

    def _generate_table_et(self, worksheet_name, table):
        table_et = ET.Element(
            "table:database-range",
            {
                "table:name": table["name"],
                "table:target-range-address": "'{}'.{}:'{}'.{}".format(
                    worksheet_name,
                    table["first_cell"],
                    worksheet_name,
                    table["last_cell"]
                ),
                "table:on-update-keep-styles": "true",
                "table:on-update-keep-size": "false"
            }
        )
        return table_et

    def _column_style_name(self, ws_index, column_number):
        return 'col_style_{}_{}'.format(ws_index, column_number)

    def _generate_column_width_defs(self, ws_index, worksheet):
        tags = []
        template = """
        <style:style style:name="{style_name}" style:family="table-column">
          <style:table-column-properties fo:break-before="auto" style:column-width="{width_cm}cm"/>
        </style:style>
        """

        for column_number, width in enumerate(worksheet.get_column_widths()):
            tags.append(template.format(
                style_name=self._column_style_name(ws_index, column_number),
                width_cm=width
            ))
   
        return "\n".join(tags)

    def _generate_content_xml(self, template_file_content):
        formats_xml = ""
        sheets_xml = ""
        tables_xml = ""
        column_width_defs_xml = ""

        for _format in self.formats.values():
            formats_xml += _format.to_xml()

        for ws_index, ws in enumerate(self.worksheets):
            worksheet = ws["worksheet"]
            sheet_et = self._generate_sheet_et(
                ws["name"], ws_index, worksheet.data_as_jagged_array(), worksheet.get_number_of_columns()
            )
            sheets_xml += ET.tostring(sheet_et, encoding="unicode")

            for table in worksheet.tables:
                table_et = self._generate_table_et(ws["name"], table)
                tables_xml += ET.tostring(table_et, encoding="unicode")

            column_width_defs_xml += self._generate_column_width_defs(ws_index, ws["worksheet"])

        return template_file_content.format(
            formats=formats_xml,
            sheets=sheets_xml,
            tables=tables_xml,
            column_width_defs=column_width_defs_xml,
        )

    def close(self):
        ods_filenames = [
            "mimetype",
            "content.xml",
            "meta.xml",
            "styles.xml",
            "META-INF/manifest.xml"
        ]
        with zipfile.ZipFile(self.filename, 'w') as zipf:
            for filename in ods_filenames:
                with open(os.path.join('ods-template', filename), 'r') as f:
                    file_content = f.read()
                if filename == "content.xml":
                    file_content = self._generate_content_xml(file_content)
                zipf.writestr(
                    filename,
                    file_content,
                    zipfile.ZIP_STORED if filename == "mimetype" else zipfile.ZIP_DEFLATED
                )
