import re

def parse_length(v):
    m = re.match(r'(-+)?(\d+)(\.\d+)?(cp|cm|emu|in|mm|pt|%)?', v)
    if m:
        sign = (-1,+1)[int(m.group(1) != '-')]
        whole = m.group(2)
        fraction = m.group(3)
        if fraction is None:
            fraction = ''
        unit = m.group(4)
        num = sign * float(whole+fraction)
        return (num, unit)
    else:
        return None

def parse_color(v):
    m = re.match(r'#[0-9a-zA-Z]{6}', v)
    if m:
        return v[1:]
    else:
        return None

def parse_font_size(v):
    # accept absolute values in pt
    return float(v)

def parse_string(v):
    return v

def parse_bool(v):
    if v.lower() == 'true':
        return True
    return False
