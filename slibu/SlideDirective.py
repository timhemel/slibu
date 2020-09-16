
from mistune.directives.base import Directive
from collections import OrderedDict
from directive_util import parse_string, parse_length


class SlideDirective(Directive):

    parsers = {
            'background': parse_string,
            'layout': parse_string,
            'left': parse_length,
            'top': parse_length,
            'width': parse_length,
            'height': parse_length,
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
            'type': 'slide',
            'raw': None,
            'params': d.items(),
        }

    def __call__(self, md):
        self.register_directive(md, 'slide')
        md.renderer.register('slide', render_ast_slide)

def render_ast_slide(children, *options):
    return {
        'type': 'slide',
        'options': dict(options),
    }

