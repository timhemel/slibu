from mistune.directives.base import Directive
from collections import OrderedDict
from .directive_util import parse_length, parse_color, parse_font_size, parse_string, parse_int_list

class SetStyleDirective(Directive):

    parsers = {
            'font_size': parse_font_size,
            'color': parse_color,
            'font_name': parse_string,
            'column_widths': parse_int_list,
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
            'type': 'setstyle',
            'raw': None,
            'params': d.items(),
        }

    def __call__(self, md):
        self.register_directive(md, 'setstyle')
        md.renderer.register('setstyle', render_ast_setstyle)


def render_ast_setstyle(children, *options):
    return {
        'type': 'setstyle',
        'options': dict(options),
    }
