"""Syntax and integrity validator for HTML, CSS, JS files."""
import sys
import re
import os
import html.parser

class HTMLValidator(html.parser.HTMLParser):
    def __init__(self):
        super().__init__()
        self.tags = []
        self.self_closing = {
            'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
            'link', 'meta', 'param', 'source', 'track', 'wbr'
        }
        self.errors = []

    def handle_starttag(self, tag, attrs):
        if tag.lower() not in self.self_closing:
            self.tags.append((tag.lower(), self.getpos()))

    def handle_endtag(self, tag):
        tag_lower = tag.lower()
        if tag_lower in self.self_closing:
            return
        if not self.tags:
            self.errors.append(f"Unexpected closing tag </{tag}> at line {self.getpos()[0]}")
            return
        last_tag, pos = self.tags.pop()
        if last_tag != tag_lower:
            self.errors.append(f"Mismatched tag: expected </{last_tag}> (opened line {pos[0]}), found </{tag}> at line {self.getpos()[0]}")

def check_html(filepath):
    print(f"Checking HTML: {filepath}")
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    validator = HTMLValidator()
    try:
        validator.feed(content)
        if validator.tags:
            for t, pos in validator.tags:
                validator.errors.append(f"Unclosed tag <{t}> opened at line {pos[0]}")
        if validator.errors:
            for err in validator.errors:
                print(f"  [ERROR] {err}")
            return False
        print("  [OK] HTML syntax valid!")
        return True
    except Exception as e:
        print(f"  [EXCEPTION] {e}")
        return False

def check_brackets_and_quotes(filepath, is_js=False):
    print(f"Checking syntax/brackets: {filepath}")
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    stack = []
    pairs = {')': '(', '}': '{', ']': '['}
    in_single_quote = False
    in_double_quote = False
    in_backtick = False
    in_line_comment = False
    in_block_comment = False

    for line_num, line in enumerate(lines, 1):
        i = 0
        in_line_comment = False
        while i < len(line):
            ch = line[i]
            prev_ch = line[i-1] if i > 0 else ''
            next_ch = line[i+1] if i + 1 < len(line) else ''

            # Comment handling
            if not in_single_quote and not in_double_quote and not in_backtick:
                if not in_block_comment and ch == '/' and next_ch == '/':
                    in_line_comment = True
                    break
                if not in_block_comment and ch == '/' and next_ch == '*':
                    in_block_comment = True
                    i += 2
                    continue
                if in_block_comment and ch == '*' and next_ch == '/':
                    in_block_comment = False
                    i += 2
                    continue

            if in_block_comment or in_line_comment:
                i += 1
                continue

            # String literal handling
            if ch == "'" and not in_double_quote and not in_backtick:
                if prev_ch != '\\':
                    in_single_quote = not in_single_quote
            elif ch == '"' and not in_single_quote and not in_backtick:
                if prev_ch != '\\':
                    in_double_quote = not in_double_quote
            elif ch == '`' and not in_single_quote and not in_double_quote and is_js:
                if prev_ch != '\\':
                    in_backtick = not in_backtick

            # Bracket checking when not inside string
            if not in_single_quote and not in_double_quote and not in_backtick:
                if ch in '({[':
                    stack.append((ch, line_num))
                elif ch in ')}]':
                    expected = pairs[ch]
                    if not stack:
                        print(f"  [ERROR] Unmatched closing '{ch}' at line {line_num}")
                        return False
                    top, open_line = stack.pop()
                    if top != expected:
                        print(f"  [ERROR] Mismatched bracket: expected '{expected}' (from line {open_line}), found '{ch}' at line {line_num}")
                        return False
            i += 1

    if in_block_comment:
        print("  [ERROR] Unclosed block comment /* ... */")
        return False
    if in_backtick:
        print("  [ERROR] Unclosed template literal ` ... `")
        return False
    if stack:
        top, open_line = stack[-1]
        print(f"  [ERROR] Unclosed bracket '{top}' opened at line {open_line}")
        return False

    print("  [OK] Brackets and syntax balance valid!")
    return True

if __name__ == '__main__':
    targets = ['index.html', 'styles.css', 'app.js']
    all_ok = True
    for t in targets:
        if os.path.exists(t):
            if t.endswith('.html'):
                if not check_html(t):
                    all_ok = False
            elif t.endswith('.css'):
                if not check_brackets_and_quotes(t, is_js=False):
                    all_ok = False
            elif t.endswith('.js'):
                if not check_brackets_and_quotes(t, is_js=True):
                    all_ok = False
    if all_ok:
        print("\nAll syntax checks passed successfully!")
    else:
        print("\nSome syntax checks failed!")
        sys.exit(1)
