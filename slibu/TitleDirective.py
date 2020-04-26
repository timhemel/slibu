from mistune.directives.base import Directive
from collections import OrderedDict
from .directive_util import parse_length, parse_color


class TitleDirective(Directive):

    parsers = {
            'left': parse_length,
            'top': parse_length,
            'width': parse_length,
            'height': parse_length,
            'bg_color': parse_color,
            'margin_left': parse_length,
            'margin_right': parse_length,
            'margin_top': parse_length,
            'margin_bottom': parse_length,
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

        return {
            'type': 'title',
            'raw': None,
            'params': d.items(),
        }

    def __call__(self, md):
        self.register_directive(md, 'title')
        md.renderer.register('title', render_ast_title)

def render_ast_title(children, *options):
    return {
        'type': 'title',
        'options': dict(options),
    }
