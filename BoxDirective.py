from mistune.directives.base import Directive
from collections import OrderedDict
from directive_util import parse_length, parse_color, parse_font_size


class BoxDirective(Directive):

    parsers = {
            'left': parse_length,
            'top': parse_length,
            'width': parse_length,
            'height': parse_length,
            'bg_color': parse_color,
            'font_size': parse_font_size,
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
            'type': 'box',
            'raw': None,
            'params': d.items(),
            'children': children
        }

    def __call__(self, md):
        self.register_directive(md, 'box')
        md.renderer.register('box', render_ast_box)

    @staticmethod
    def parse_text(m):
        text = m.group('text')
        if not text.strip():
            return ''
        leading = len(m.group(1)) + 2

        text = '\n'.join([ l[leading:] for l in text.splitlines() ]).lstrip('\n') + '\n'
        return text


def render_ast_box(children, *options):
    return {
        'type': 'box',
        'children': children,
        'options': dict(options),
    }
