from collections import deque
import html
import json
import re

from flask import Flask, render_template, request

from docx_equation.omml import omml2tex

escape = html.escape

app = Flask(__name__)


def convert_ms_equation(omath):
    def add_mt(match):
        content = re.sub(
            r'(|>)([^<>]+)(<|$)', lambda m: m.group(1) + '<m:t>{}</m:t>'.
            format(m.group(2)) + m.group(3), match.group(1))
        return '<m:r>{}</m:r>'.format(content)

    for tag in ['span', 'i']:
        omath = re.sub(r'</?{}[^<>]*>'.format(tag), '', omath)
    omath = re.sub(r'<m:r>(.+?)</m:r>', add_mt, omath, flags=re.S)
    latex = omml2tex(omath)
    latex = html.escape(latex)
    return '<span class="latex">{}</span>'.format(latex)


def handle_paste(content):
    content = re.sub(r'\s+', ' ', content)
    # delete no use tag
    for tag in ['meta', 'link']:
        pattern = re.compile(r'<{}\s*.+?>'.format(tag), flags=re.S)
        content = pattern.sub('', content)
    for tag in ['style']:
        pattern = re.compile(r'<{}[^<>]*?>.+?</{}>'.format(tag, tag),
                             flags=re.S)
        content = pattern.sub('', content)
    # delete inline style
    pattern = re.compile(r'(?P<tag_begin><[a-z][0-9a-zA-Z][^>]*)'
                         r'\sstyle\s*=\s*(?P<quote>[\'"])[^>]*?(?P=quote)')
    content = pattern.sub(lambda m: m.group('tag_begin'), content)
    # handle comment
    condition_pattern = re.compile(r'((?:<!--[^<>]+>)|(?:<!\[endif\]-->))')
    condition_stack = deque()
    lst = condition_pattern.split(content)
    content = ''
    for item in lst:
        if '<!--[if ' in item:
            condition_stack.append(item)
        elif item == '<![endif]-->':
            condition_stack.pop()
        else:
            if condition_stack:
                if '[if gte msEquation' in condition_stack[-1]:
                    content += convert_ms_equation(item)
            else:
                content += item
    content = re.sub(r'<!--.+?-->', '', content, flags=re.S)
    return content


@app.route('/')
def index():
    return render_template('demo.html')


@app.route('/api', methods=['POST'])
def api():
    data = json.loads(request.data.decode())
    content = data['content']
    content = handle_paste(content)
    return {'content': content}
