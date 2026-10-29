import { createRequire } from 'node:module';
import fs from 'node:fs/promises';
import path from 'node:path';

function options(args) {
  const result = {};
  for (let i = 0; i < args.length; i += 2) {
    if (!['--scene', '--out', '--modules'].includes(args[i]) || !args[i + 1]) throw new Error('用法：export.mjs --scene scene.json --out result.pptx [--modules node_modules]');
    result[args[i].slice(2)] = args[i + 1];
  }
  if (!result.scene || !result.out) throw new Error('必须指定 --scene 和 --out。');
  return result;
}

function validate(deck) {
  if (!deck.title || !deck.theme || !Array.isArray(deck.slides)) throw new Error('场景必须包含 title、theme 和 slides。');
  if (!Number.isInteger(deck.brief?.pageCount) || deck.brief.pageCount < 1 || deck.slides.length !== deck.brief.pageCount) throw new Error('brief.pageCount 必须为已明确的总页数，并与实际幻灯片数相同。');
  const ids = new Set();
  for (const slide of deck.slides) {
    if (!slide.id || !Array.isArray(slide.elements)) throw new Error('页面必须包含 id 和 elements。');
    for (const e of slide.elements) {
      if (!e.id || ids.has(e.id)) throw new Error('对象 id 必须唯一。');
      ids.add(e.id);
      if (!['text', 'shape', 'image'].includes(e.type) || !['x', 'y', 'w', 'h'].every(k => Number.isFinite(e[k])) || e.w <= 0 || e.h <= 0) throw new Error(`对象 ${e.id} 类型或尺寸无效。`);
      if (e.type === 'text' && typeof e.text !== 'string') throw new Error(`对象 ${e.id} 缺少 text。`);
      if (e.type === 'image' && !e.path) throw new Error(`对象 ${e.id} 缺少图片 path。`);
      if (e.type === 'shape' && !['rect', 'roundRect', 'ellipse', 'line'].includes(e.shape)) throw new Error(`对象 ${e.id} 形状无效。`);
      if (e.opacity !== undefined && (!Number.isFinite(e.opacity) || e.opacity < 0 || e.opacity > 1)) throw new Error(`对象 ${e.id} 透明度无效。`);
      for (const k of ['fill', 'stroke', 'color']) if (e[k] && !(k === 'fill' && e[k] === 'none') && !/^#[0-9A-Fa-f]{6}$/.test(e[k])) throw new Error(`对象 ${e.id} 的 ${k} 需要 #RRGGBB。`);
    }
  }
}

async function main() {
  const args = options(process.argv.slice(2));
  const scenePath = path.resolve(args.scene), output = path.resolve(args.out);
  const deck = JSON.parse(await fs.readFile(scenePath, 'utf8'));
  validate(deck);
  const require = args.modules ? createRequire(path.join(path.resolve(args.modules), '__pptmaker__.cjs')) : createRequire(import.meta.url);
  const PptxGenJS = require('pptxgenjs');
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'SCENE', width: 13.3333333333, height: 7.5 });
  pptx.layout = 'SCENE'; pptx.author = 'Arknights PPTmaker'; pptx.title = deck.title; pptx.subject = deck.theme;
  const color = (value, fallback = '#242825') => (value || fallback).replace('#', '');
  const counts = { textObjects: 0, shapeObjects: 0, pictureObjects: 0 };
  for (const page of deck.slides) {
    const slide = pptx.addSlide();
    slide.background = { color: color(page.background, '#F2EDD8') };
    const sourceIds = new Set(page.sourceIds || []);
    for (const e of page.elements) {
      (e.sourceIds || []).forEach(id => sourceIds.add(id));
      const box = { x: e.x / 96, y: e.y / 96, w: e.w / 96, h: e.h / 96, rotate: e.rotation || 0, objectName: `${e.name || e.id} ${e.id}` };
      const transparency = Math.round((1 - (e.opacity ?? 1)) * 100);
      if (e.type === 'text') {
        slide.addText(e.text, { ...box, lang: 'zh-CN', fontFace: e.fontFace || 'Microsoft YaHei', fontSize: (e.fontSize || 28) * .75, color: color(e.color), bold: !!e.bold, italic: !!e.italic, align: e.align || 'left', valign: 'top', margin: 0, paraSpaceAfter: 0, lineSpacingMultiple: e.lineHeight || 1.35, transparency });
        counts.textObjects++;
      } else if (e.type === 'shape') {
        slide.addShape(pptx.ShapeType[e.shape], { ...box, fill: { color: color(e.fill === 'none' ? '#FFFFFF' : e.fill), transparency: e.fill === 'none' ? 100 : transparency }, line: { color: color(e.stroke || (e.fill === 'none' ? undefined : e.fill)), transparency: e.stroke ? transparency : 100, width: (e.strokeWidth || 1) * .75 } });
        counts.shapeObjects++;
      } else {
        const imagePath = path.resolve(path.dirname(scenePath), e.path);
        const image = { ...box, sizing: { type: e.fit || 'contain', w: box.w, h: box.h }, altText: e.name || e.id, transparency };
        if (path.extname(imagePath).toLowerCase() === '.webp') {
          const png = await require('sharp')(imagePath, { limitInputPixels: 40_000_000 }).rotate().png().toBuffer();
          image.data = `image/png;base64,${png.toString('base64')}`;
        } else image.path = imagePath;
        slide.addImage(image); counts.pictureObjects++;
      }
    }
    slide.addNotes(`活动主题：${deck.theme}\n页面：${page.title || page.id}\n资料块或用户回答：${[...sourceIds].join('、')}\n${page.notes || ''}\n重新排版的文字为原生对象，装饰与图片独立分层。`);
  }
  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  await fs.mkdir(path.dirname(output), { recursive: true });
  await fs.writeFile(output, buffer);
  process.stdout.write(JSON.stringify({ output, pageCount: deck.slides.length, ...counts }) + '\n');
}

main().catch(error => { process.stderr.write(`${error.message}\n`); process.exitCode = 1; });
