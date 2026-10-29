"""Read local sources or inspect actual PPTX objects; no server or third-party Python packages."""
import argparse
import json
import mimetypes
import posixpath
import sys
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


def text_of(element):
    parts = []
    for node in element.iter():
        if node.tag == '{%s}t' % NS['w']:
            parts.append(node.text or '')
        elif node.tag in ('{%s}br' % NS['w'], '{%s}cr' % NS['w']):
            parts.append('\n')
        elif node.tag == '{%s}tab' % NS['w']:
            parts.append('\t')
    return ''.join(parts).strip()


def extract(files, output):
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    asset_dir = output.parent / (output.stem + '-assets-' + uuid.uuid4().hex[:8])
    bundle = {'files': [], 'blocks': [], 'assets': [], 'warnings': []}

    def add(kind, text, file, location, **extra):
        block = {'id': 'f%d-b%d' % (len(bundle['files']), len(bundle['blocks']) + 1),
                 'kind': kind, 'text': text,
                 'source': {'fileName': file.name, 'location': location}, **extra}
        bundle['blocks'].append(block)

    def asset(data, name, suffix):
        asset_dir.mkdir(parents=True, exist_ok=True)
        key = 'img-' + uuid.uuid4().hex[:12]
        target = asset_dir / (key + suffix)
        target.write_bytes(data)
        bundle['assets'].append({'id': key, 'name': name,
                                 'mime': mimetypes.guess_type(target.name)[0] or 'application/octet-stream',
                                 'path': target.relative_to(output.parent).as_posix()})
        return key

    for input_path in files:
        file = Path(input_path).resolve()
        if file.stat().st_size > 30 * 1024 * 1024:
            raise ValueError('%s 超过 30 MB，请分拆资料。' % file.name)
        bundle['files'].append(file.name)
        suffix = file.suffix.lower()
        if suffix in ('.png', '.jpg', '.jpeg', '.webp'):
            key = asset(file.read_bytes(), file.name, suffix)
            add('image', file.stem, file, '上传图片', assetId=key)
        elif suffix in ('.txt', '.md'):
            for i, paragraph in enumerate(file.read_text(encoding='utf-8-sig').split('\n\n'), 1):
                if paragraph.strip():
                    add('paragraph', paragraph.strip(), file, '文本段落 %d' % i)
        elif suffix == '.docx':
            with zipfile.ZipFile(file) as archive:
                if len(archive.infolist()) > 3000 or sum(x.file_size for x in archive.infolist()) > 120 * 1024 * 1024:
                    raise ValueError('%s 解压后过大。' % file.name)
                body = ET.fromstring(archive.read('word/document.xml')).find('w:body', NS)
                if body is None:
                    raise ValueError('%s 缺少正文。' % file.name)
                rels = {}
                rel_path = 'word/_rels/document.xml.rels'
                if rel_path in archive.namelist():
                    rels = {r.attrib['Id']: r.attrib for r in ET.fromstring(archive.read(rel_path))}
                images = {}
                paragraph_count = table_count = 0
                for element in body:
                    if element.tag == '{%s}p' % NS['w']:
                        paragraph_count += 1
                        location = '段落 %d' % paragraph_count
                        value = text_of(element)
                        style = element.find('w:pPr/w:pStyle', NS)
                        style_name = style.attrib.get('{%s}val' % NS['w'], '') if style is not None else ''
                        if value:
                            add('heading' if style_name.lower().startswith(('heading', '标题')) else 'paragraph', value, file, location)
                    elif element.tag == '{%s}tbl' % NS['w']:
                        table_count += 1
                        location = '表格 %d' % table_count
                        rows = [['\n'.join(text_of(p) for p in cell.findall('w:p', NS))
                                 for cell in row.findall('w:tc', NS)] for row in element.findall('w:tr', NS)]
                        add('table', '\n'.join(' / '.join(row) for row in rows), file, location, rows=rows)
                    else:
                        continue
                    for blip in element.findall('.//a:blip', NS):
                        rid = blip.attrib.get('{%s}embed' % NS['r'])
                        relation = rels.get(rid, {})
                        if not rid or relation.get('TargetMode') == 'External':
                            bundle['warnings'].append('%s %s：外链图片未读取。' % (file.name, location))
                            continue
                        if rid not in images:
                            target = posixpath.normpath(posixpath.join('word', relation.get('Target', '')))
                            if relation.get('Target', '').startswith('/'):
                                target = relation['Target'].lstrip('/')
                            if target not in archive.namelist():
                                raise ValueError('%s %s：找不到内嵌图片。' % (file.name, location))
                            images[rid] = asset(archive.read(target), posixpath.basename(target), Path(target).suffix.lower())
                        add('image', '文档图片', file, location + ' 内嵌图片', assetId=images[rid])
                bundle['warnings'].append('%s：已读取主正文；页眉、页脚、脚注、批注、复杂文本框或旧式 VML 图像需要另行核对。' % file.name)
        else:
            raise ValueError('%s：请通过文档/PDF工具读取此格式。' % file.name)
    if not bundle['blocks']:
        raise ValueError('没有读取到可用资料。')
    output.write_text(json.dumps(bundle, ensure_ascii=False, indent=2), encoding='utf-8')
    return {'output': str(output), 'blocks': len(bundle['blocks']), 'images': len(bundle['assets']), 'warnings': bundle['warnings']}


def inspect_pptx(file):
    pages = []
    with zipfile.ZipFile(file) as archive:
        names = [n for n in archive.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
        names.sort(key=lambda n: int(n.rsplit('slide', 1)[1].split('.')[0]))
        for name in names:
            root = ET.fromstring(archive.read(name))
            shapes = root.findall('.//p:sp', NS)
            texts = [s for s in shapes if s.findall('.//a:t', NS)]
            pages.append({'slide': len(pages) + 1, 'textObjects': len(texts),
                          'shapeObjects': len(shapes) - len(texts),
                          'pictureObjects': len(root.findall('.//p:pic', NS)),
                          'text': [''.join(s.itertext()) for s in texts]})
    return {'pageCount': len(pages), **{k: sum(p[k] for p in pages) for k in ('textObjects', 'shapeObjects', 'pictureObjects')}, 'pages': pages}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    extract_parser = commands.add_parser('extract')
    extract_parser.add_argument('--out', required=True)
    extract_parser.add_argument('files', nargs='+')
    inspect_parser = commands.add_parser('inspect')
    inspect_parser.add_argument('file')
    args = parser.parse_args()
    try:
        result = extract(args.files, args.out) if args.command == 'extract' else inspect_pptx(args.file)
        print(json.dumps(result, ensure_ascii=True, indent=2))
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as error:
        print(str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
