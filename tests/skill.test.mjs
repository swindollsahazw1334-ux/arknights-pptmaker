import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import JSZip from 'jszip';
import { makeDocx } from './helpers/docx.mjs';

const root = process.env.PPTMAKER_SKILL_ROOT || path.resolve('.');
const modules = path.resolve('node_modules');

test('standalone skill extracts DOCX sources, numeric boundaries and embedded pictures without a server', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'ppt-skill-source-'));
  const zip = await JSZip.loadAsync(await makeDocx());
  const xml = await zip.file('word/document.xml').async('string');
  zip.file('word/document.xml', xml.replace('活动预算为1200元。', '活动预算为1200元。</w:t><w:br/><w:t>12</w:t><w:br/><w:t>34'));
  const file = path.join(dir, '用户资料 空格.docx');
  await writeFile(file, await zip.generateAsync({ type: 'nodebuffer' }));
  const output = path.join(dir, 'source.json');
  const run = spawnSync('python', [path.join(root, 'scripts/sources.py'), 'extract', '--out', output, file], { cwd: os.tmpdir(), encoding: 'utf8' });
  assert.equal(run.status, 0, run.stderr);
  const source = JSON.parse(await readFile(output, 'utf8'));
  assert(source.blocks.some((b) => b.kind === 'heading' && b.text === '海岸清洁方案'));
  assert(source.blocks.some((b) => b.text.includes('12\n34')));
  assert(source.blocks.some((b) => b.kind === 'table' && b.rows[1][1] === '1200元'));
  const image = source.assets[0];
  assert((await readFile(path.resolve(dir, image.path))).length > 0);
  assert(source.blocks.some((b) => b.kind === 'image' && b.assetId === image.id && b.source.fileName === '用户资料 空格.docx'));
});

function scene() {
  return { title: '对话生成稿', theme: '孤星', brief: { pageCount: 2, density: 'balanced' }, slides: [
    { id: 's1', title: '封面', background: '#F2EDD8', sourceIds: ['source-1'], elements: [
      { id: 'title', name: '封面标题', type: 'text', x: 80, y: 120, w: 900, h: 90, text: '活动预算 1200元', fontSize: 48, color: '#242825' },
      { id: 'band', name: '独立橙色装饰', type: 'shape', shape: 'rect', x: 50, y: 60, w: 12, h: 520, fill: '#E8753C' },
    ] },
    { id: 's2', title: '资料与配图', background: '#F2EDD8', sourceIds: ['source-2'], elements: [
      { id: 'caption', name: '图片说明', type: 'text', x: 80, y: 80, w: 600, h: 70, text: '原始照片与独立图层', fontSize: 32 },
      { id: 'image', name: '资料图片', type: 'image', x: 100, y: 200, w: 320, h: 240, path: 'photo.png', fit: 'contain' },
    ] },
  ] };
}

test('standalone skill exports editable text, movable decoration, image and source notes from a scene', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'ppt-skill-export-'));
  const fixture = await JSZip.loadAsync(await makeDocx());
  await writeFile(path.join(dir, 'photo.png'), await fixture.file('word/media/image1.png').async('nodebuffer'));
  const json = path.join(dir, 'scene.json'), pptx = path.join(dir, '对话稿.pptx');
  await writeFile(json, JSON.stringify(scene()));
  const run = spawnSync(process.execPath, [path.join(root, 'scripts/export.mjs'), '--modules', modules, '--scene', json, '--out', pptx], { cwd: os.tmpdir(), encoding: 'utf8' });
  assert.equal(run.status, 0, run.stderr);
  const zip = await JSZip.loadAsync(await readFile(pptx));
  const cover = await zip.file('ppt/slides/slide1.xml').async('string');
  const content = await zip.file('ppt/slides/slide2.xml').async('string');
  assert(cover.includes('<a:t>活动预算 1200元</a:t>'));
  assert(cover.includes('独立橙色装饰'));
  assert.equal((cover.match(/<p:sp>/g) || []).length, 2);
  assert.equal((content.match(/<p:pic>/g) || []).length, 1);
  assert((await zip.file('ppt/notesSlides/notesSlide1.xml').async('string')).includes('source-1'));
  const inspect = spawnSync('python', [path.join(root, 'scripts/sources.py'), 'inspect', pptx], { encoding: 'utf8' });
  assert.equal(inspect.status, 0, inspect.stderr);
  const report = JSON.parse(inspect.stdout);
  assert.equal(report.pageCount, 2);
  assert.equal(report.textObjects, 2);
  assert.equal(report.shapeObjects, 1);
  assert.equal(report.pictureObjects, 1);
});

test('skill refuses an exported page count that differs from the conversational brief', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'ppt-skill-count-'));
  const value = scene(); value.brief.pageCount = 3;
  const json = path.join(dir, 'scene.json');
  await writeFile(json, JSON.stringify(value));
  const run = spawnSync(process.execPath, [path.join(root, 'scripts/export.mjs'), '--modules', modules, '--scene', json, '--out', path.join(dir, 'bad.pptx')], { encoding: 'utf8' });
  assert.notEqual(run.status, 0);
  assert.match(run.stderr, /pageCount/);
});
