
from mistune.directives.base import Directive
from collections import OrderedDict
from .directive_util import parse_length, parse_color, parse_font_size, parse_string

class TextStyleDirective(Directive):

    parsers = {
            'font_size': parse_font_size,
            'color': parse_color,
            'font_name': parse_string,
    }


    def parse(self, block, m, state):
        arg = m.group('value')
        options = self.parse_options(m)
        d = OrderedDict()
        for k,v in options:
            try:
                d[k] = self.parsers[k](v)
            except KeyError:
                pass

        text = self.parse_text(m)
        rules = list(block.rules)
        rules.remove('directive')
        children = block.parse(text, state, rules)
        return {
            'type': 'textstyle',
            'raw': None,
            'params': d.items(),
            'children': children
        }

    def __call__(self, md):
        self.register_directive(md, 'textstyle')
        md.renderer.register('textstyle', render_ast_textstyle)

    @staticmethod
    def parse_text(m):
        text = m.group('text')
        if not text.strip():
            return ''
        leading = len(m.group(1)) + 2

        text = '\n'.join([ l[leading:] for l in text.splitlines() ]).lstrip('\n') + '\n'
        return text


def render_ast_textstyle(children, *options):
    return {
        'type': 'textstyle',
        'children': children,
        'options': dict(options),
    }
