
from pygments import highlight
from pygments.lexers import get_lexer_by_name
from pygments.formatter import Formatter
import re
from presentation import PPTXPresentation

class PresentationPygmentsFormatter(Formatter):

    def __init__(self, presentation, **options):
        Formatter.__init__(self, **options)
        self.presentation = presentation

    def format(self, tokensource, outfile):

        def get_font_attributes(ttype):
            d = {
                'italic': None,
                'bold': None,
                'underline': None,
                'fontname': None,
                'color': None,
                'border': None,
                'bg_color': None,
            }

            style = self.style.styles[ttype].split()
            for attr in style:
                # bold: render text as bold
                if attr == 'bold':
                    d['bold'] = True
                elif attr == 'nobold':
                    d['bold'] = False
                # italic: render text italic
                elif attr == 'italic':
                    d['italic'] = True
                elif attr == 'noitalic':
                    d['italic'] = False
                # underline render text underlined
                # nounderline don’t render text underlined
                elif attr == 'underline':
                    d['underline'] = True
                elif attr == 'nounderline':
                    d['underline'] = False
                # bg: transparent background
                # bg:#000000 background color (black)
                elif attr == 'bg:':
                    # background not supported on text run, 
                    d['bg_color'] = None
                    pass
                elif re.match(r'^bg:#[0-9A-Za-z]{6}', attr):
                    # pptx does not support background fill for text runs,
                    # only for complete shapes.
                    # rgb = RGBColor.from_string(attr[4:])
                    # run.font.fill.back_color.rgb = rgb
                    d['bg_color'] = attr[4:]
                # border: no border
                elif re.match(r'^border', attr):
                    d['border'] = None
                # border:#ffffff border color (white)
                elif re.match(r'^border:#[0-9A-Za-z]{6}', attr):
                    d['border'] = attr[8:]
                # #ff0000 text color (red)
                elif re.match(r'^#[0-9A-Za-z]{6}', attr) is not None:
                    # rgb = RGBColor.from_string(attr[1:])
                    # run.font.color.rgb = rgb
                    d['color'] = attr[1:]
                # noinherit don’t inherit styles from supertoken

            return d
 
        elements = [ (text, get_font_attributes(ttype)) for ttype, text in tokensource ]

        if elements[-1][0].rstrip('\n') == '':
            del elements[-1]

        for text, style in elements:
            self.presentation.add_text(text, verbatim=True, **style)

class PPTXWriter:

    def __init__(self):
        self.presentation = PPTXPresentation()

    def save(self, outfn):
        self.presentation.save(outfn)

    def feed(self, ast):
        for e in ast:
            print("DEBUG:", "e=", e)
            try:
                h = self.handlers[e['type']]
                try:
                    h(self, e)
                except Exception as e:
                    print("EXCEPTION", e)
                    raise(e)
            except KeyError:
                self.handle_undefined(e)

    def handle_undefined(self, e):
        print("ERROR: No handler for", e['type'])

    def handle_heading(self, e):
        if e['level'] == 1:
            self.presentation.add_slide(layout="1_Section Header", title=e['children'][0]['text'])
        elif e['level'] == 2:
            self.presentation.add_slide(layout="Title and Content", title=e['children'][0]['text'])

    def handle_paragraph(self, e):
        if self.presentation.has_content_placeholder():
            self.presentation.add_paragraph()
            self.feed(e['children'])

    def handle_block_text(self, e):
        if self.presentation.has_content_placeholder():
            self.presentation.add_paragraph()
            self.feed(e['children'])

    def handle_text(self, e):
        self.presentation.add_text(e['text'])

    def handle_emphasis(self, e):
        text = "".join([ x.get('text') for x in e['children'] ])
        self.presentation.add_text(text, italic=True)

    def handle_strong(self, e):
        text = "".join([ x.get('text') for x in e['children'] ])
        self.presentation.add_text(text, bold=True)

    def handle_codespan(self, e):
        self.presentation.add_text(e['text'], verbatim=True, fontname='Courier')

    def handle_block_quote(self, e):
        # TODO: fix fontname
        self.presentation.push_text_layout(italic=True, fontname='Comic Sans MS')
        self.feed(e['children'])
        self.presentation.pop_text_layout()

    def handle_list(self, e):
        self.presentation.push_text_layout(list_order=e['ordered'])
        self.feed(e['children'])
        self.presentation.pop_text_layout()

    def handle_list_item(self, e):
        self.presentation.push_text_layout(indent_level=e['level'], show_bullet=True)
        self.feed(e['children'])
        self.presentation.pop_text_layout()

    def handle_image(self, e):
        self.presentation.add_picture(e['src'])

    def handle_block_code(self, e):
        if self.presentation.has_content_placeholder():
            self.presentation.push_text_layout(fontname='Courier')
            self.presentation.add_paragraph()
            text = e['text'].rstrip('\n')
            if e['info'] is None:
                self.presentation.add_text(text, verbatim=True)
            else:
                lexer = get_lexer_by_name(e['info'].strip())
                formatter = PresentationPygmentsFormatter(self.presentation)
                result = highlight(text, lexer, formatter)

            self.presentation.pop_text_layout()

    def handle_table(self, e):
        rows = len(e['children'][1]['children']) + 1
        cols = len(e['children'][0]['children'])
        self.presentation.add_table(rows, cols)
        self.row = 0
        self.feed(e['children'])

    def handle_table_head(self, e):
        self.col = 0
        for ce in e['children']:
            self.feed([ce])
            self.col += 1
        self.row += 1

    def handle_table_body(self, e):
        self.feed(e['children'])

    def handle_table_row(self, e):
        self.col = 0
        for ce in e['children']:
            self.feed([ce])
            self.col += 1
        self.row += 1

    def handle_table_cell(self, e):
        self.presentation.start_table_cell(self.row, self.col)
        self.feed(e['children'])
        self.presentation.end_table_cell()

    def handle_box(self, e):
        self.presentation.start_box(**e['options'])
        self.feed(e['children'])
        self.presentation.end_box()

    def handle_img(self, e):
        self.presentation.add_picture(e['src'], **e['options'])

    def handle_title(self, e):
        self.presentation.set_title_box(e['options'])

    def handle_textstyle(self, e):
        self.presentation.push_text_layout(**e['options'])
        self.feed(e['children'])
        self.presentation.pop_text_layout()

    def handle_setstyle(self, e):
        self.presentation.set_text_layout(**e['options'])

    def handle_linebreak(self, e):
        if self.presentation.has_content_placeholder():
            self.presentation.add_paragraph()

    def handle_slide(self, e):
        layout = e['options'].get('layout')
        if layout is not None:
            self.presentation.set_slide_layout(layout)

        background = e['options'].get('background')
        if background is not None:
            self.presentation.set_slide_background(background, **e['options'])

    # TODO:
    # source code from file
    handlers = {
        'heading': handle_heading,
        'paragraph': handle_paragraph,
        'text': handle_text,
        'emphasis': handle_emphasis,
        'strong': handle_strong,
        'image': handle_image,
        'list': handle_list,
        'list_item': handle_list_item,
        'block_text': handle_block_text,
        'img': handle_img,
        'box': handle_box,
        'table': handle_table,
        'table_head': handle_table_head,
        'table_body': handle_table_body,
        'table_row': handle_table_row,
        'table_cell': handle_table_cell,
        'block_code': handle_block_code,
        'codespan': handle_codespan,
        'block_quote': handle_block_quote,
        'title': handle_title,
        'textstyle': handle_textstyle,
        'setstyle': handle_setstyle,
        'linebreak': handle_linebreak,
        'slide': handle_slide,
    }



