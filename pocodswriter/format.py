import xml.etree.cElementTree as ET

class Format:
    def __init__(self, *args):
        self.props = args[0] if args else None
        self.name = args[1] if args and len(args) > 0 else None

    def get_frozen_props(self):
        return tuple(sorted(self.props.items()))

    def __repr__(self):
        return 'Format({}, {})'.format(repr(self.props), repr(self.name))

    def to_xml(self):
        font_name = self.props["font_name"] or "Arial"
        font_size = self.props["font_size"] or "12pt"
        bold = "bold" in self.props and self.props["bold"]
        bold_template = 'fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"'
        bottom_border = 'bottom' in self.props and self.props['bottom'] == 1
        bottom_border_template = '<style:table-cell-properties fo:border-bottom="0.74pt solid #000000" fo:border-left="none" fo:border-right="none" fo:border-top="none"/>'
        xml_template = """
        <style:style style:name="{format_name}" style:family="table-cell" style:parent-style-name="Default">
          <style:text-properties
            style:font-name="{font_name}"
            fo:font-size="{font_size}"
            style:font-size-asian="{font_size}"
            style:font-size-complex="{font_size}"
            {bold_stuff}/>
            {bottom_border_stuff}
        </style:style>
        """
        return xml_template.format(**{
            "font_name": font_name,
            "font_size": font_size,
            "format_name": self.name,
            "bold_stuff": bold_template if bold else "",
            "bottom_border_stuff": bottom_border_template if bottom_border else ""
        })
