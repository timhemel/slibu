from mistune.directives.base import Directive
from collections import OrderedDict
from directive_util import parse_length

class ImageDirective(Directive):

    parsers = {
            'left': parse_length,
            'top': parse_length,
            'width': parse_length,
            'height': parse_length
    }

    def parse(self, block, m, state):
        d = OrderedDict()
        d['src'] = m.group(3).strip()
        for k in ['left','top','width','height']:
            d[k] = None
        options = self.parse_options(m)
        for k,v in options:
            try:
                d[k] = self.parsers[k](v)
            except KeyError:
                pass

        # return {'type': 'img', 'raw': None }
        return {'type': 'img', 'raw': None, 'params': list(d.values()) }

    def __call__(self, md):
        self.register_directive(md, 'img')
        md.renderer.register('img', render_ast_image)


def render_ast_image(text, src, left, top, width, height):
    return {
        'type': 'img',
        'text': text,
        'src': src,
        'options': {
            'left': left,
            'top': top,
            'width': width,
            'height': height
        }
    }
    
