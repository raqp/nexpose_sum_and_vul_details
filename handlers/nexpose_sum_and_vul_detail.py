from docx import Document
from docx.document import Document as _Document
from docx.shared import Pt, RGBColor
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table, _Row
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime


class ExecutiveSummary:

    def __init__(self, config, source_file, destination_path, input_document_type, os_platform):
        if input_document_type.lower() == 'sum':
            config = config['executive_summary']
            self.mapper = {
                'IP Address': 'Component',
                'Vulnerabilities Noted per IP address': 'Vulnerabilities Noted per Component',
                'Severity Level': 'Severity Level',
                'CVSSv2 Score': 'CVSS Score',
                'Compliance Status': 'Compliance Status',
                'Exceptions, False Positives, or Compensating Controls Noted by the ASV for this Vulnerability':
                    'Exceptions, False Positives, or Compensating Controls Noted by the ASV for this Vulnerability',
                'Scan Customer Company: ': 'Scan Customer Company: ',
                'ASV Company: ': 'ASV Company: ',
                'Remediation Step': 'Remediation Step',
                'Estimated Time': 'Estimated Time'
            }
        elif input_document_type.lower() == 'vul':
            config = config['vulnerability']
            self.mapper = {
                'Scan Customer Company:': 'Scan Customer Company: ',
                'ASV Company:': 'ASV Company: ',
                'Severity': 'Severity',
                'High': 'High',
                'IP Address': 'Component',
                'Port': 'Detected Open Ports, Services/ Protocols',
                'Evidence': 'Vulnerability',
                'Compliance Status': 'Compliance Status',
                'Exceptions, False Positives, or Compensating Controls Noted by the ASV for this Vulnerability':
                    'Details',
                'Instance': 'Instance'
            }
        self.input_document_type = input_document_type
        # self.document = Document(r'C:\Users\Acer\Desktop\work\nexpose_vulnerability\NExpose\nes.docx')
        self.document = Document(source_file)
        self.destination_path = destination_path
        self.slash = "/"
        if os_platform == 'windows':
            self.slash = "\\"
        self.destination = destination_path if destination_path.endswith(self.slash) else destination_path + self.slash
        self.file_name = f'report_{datetime.now().strftime("%d-%b-%Y-%H-%M")}.docx'
        self.paragraphs = self.document.paragraphs
        self.tables = self.document.tables
        self.font_name = config['font_name']
        self.table_header_font_size = config['table_header_font_size']
        self.page_height = config['page_height']
        self.page_width = config['page_width']
        self.default_color = config['default_color']
        self.p_18_font_size = config['p_18_font_size']
        self.p_12_font_size = config['p_12_font_size']
        self.p_10_font_size = config['p_10_font_size']
        self.p_10_font_color = config['p_10_font_color']
        self.table_header_color = config['table_header_color']
        self.table_content_font_size = config['table_content_font_size']

    def start(self):
        self.iterate()
        self.delete_useful_tables()
        self.change_tables()
        self.save_document()

    @staticmethod
    def iter_block_items(parent):
        if isinstance(parent, _Document):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        elif isinstance(parent, _Row):
            parent_elm = parent._tr
        else:
            raise ValueError("something's not right")
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def delete_useful_tables(self):
        for active_table in self.tables:
            if active_table.cell(0, 0).paragraphs[0].text == '':
                active_table._element.getparent().remove(active_table._element)

    @staticmethod
    def set_object_color(obj, rgb_color):
        if rgb_color:
            obj.font.color.rgb = RGBColor(*rgb_color)

    @staticmethod
    def set_paragraph_font_size(run, font_size, bold=None):
        run.font.size = Pt(font_size)
        if bold is None:
            run.bold = bold

    def iterate(self):
        for block in self.iter_block_items(self.document):
            if isinstance(block, Paragraph):
                font_size = 9
                color = None
                bold = None
                if block.text.startswith('Part 2a') or block.text.startswith('2'):
                    block.text = ''
                elif block.text.startswith('Part 2b'):
                    index = block.text.index('.')
                    text = f'Part 2{block.text[index:]}'
                    block.text = text
                if block.text:
                    if block.style.font.size == Pt(18):
                        if self.input_document_type == 'vul':
                            block.text = self.format_paragraph_text(block.text)
                        color = self.default_color
                        font_size = self.p_18_font_size
                    elif block.style.font.size == Pt(12):
                        color = self.default_color
                        font_size = self.p_12_font_size
                    elif block.style.font.size == Pt(10):
                        color = self.p_10_font_color
                        font_size = self.p_10_font_size
                    self.set_object_color(block.runs[0], color)
                    self.set_paragraph_font_size(block.runs[0], font_size, bold)

    @staticmethod
    def format_paragraph_text(item):
        text = item
        if item[0].isdigit():
            index = item.find(' ')
            text = f'{item[index + 1:]}'
        return text

    def change_tables(self):
        for index, table in enumerate(self.tables):
            borders = ['left', 'right', 'top']
            for cell in table.rows[0].cells:
                self.set_table_header_bg_color(cell)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.text in self.mapper:
                            run.text = self.mapper[run.text]
                            self.set_object_color(run, self.default_color)
            if index == 1 and self.input_document_type != 'vul':
                borders.append('bottom')
            self.set_table_styling(table, *borders)

    @staticmethod
    def set_table_header_bg_color(cell):
        tc = cell._tc
        tbl_cell_properties = tc.get_or_add_tcPr()
        cl_shading = OxmlElement('w:shd')
        cl_shading.set(qn('w:fill'), "00000")  # Hex of Dark Blue Shade {R:0x00, G:0x51, B:0x9E}
        tbl_cell_properties.append(cl_shading)
        return cell

    @staticmethod
    def set_table_styling(table, *args):
        tbl = table._tbl
        cell_number = 0
        coll_count = len(table.columns)
        x = ['left', 'right', 'top', 'bottom']
        for cell in tbl.iter_tcs():
            tc_pr = cell.tcPr
            tc_borders = OxmlElement("w:tcBorders")
            for border in args:
                side = OxmlElement(f'w:{border}')
                side.set(qn("w:val"), "nil")
                tc_borders.append(side)
            for i in set(x).difference(args):
                side = OxmlElement(f'w:{i}')
                side.set(qn("w:val"), "single")
                if cell_number < coll_count:
                    side.set(qn("w:sz"), "12")
                    cell_number += 1
                    side.set(qn("w:color"), "4f2d7f")
                else:
                    side.set(qn("w:sz"), "5")
                    side.set(qn("w:color"), "b5b5b5")
                tc_borders.append(side)
            tc_pr.append(tc_borders)

    def save_document(self):
        self.document.save(f'{self.destination_path}{self.file_name}')


