#!/usr/bin/env node
/**
 * scripts/bump-version.js
 *
 * 목적: package.json 버전을 읽어
 *   1) gas/*.gs 파일 상단 주석의 버전 문자열 일괄 치환
 *   2) src/api.ts 내 APP_VERSION 상수 치환
 *
 * 실행: node scripts/bump-version.js
 * 연동: package.json "prebuild" 훅에서 자동 실행
 *
 * NOTE: package.json "type": "module" 환경이므로 ESM(import) 문법 사용
 */

import { readFileSync, writeFileSync, readdirSync, existsSync } from 'fs';
import { resolve, join, dirname } from 'path';
import { fileURLToPath } from 'url';

// ─────────────────────────────────────────────
// 경로 설정 (ESM __dirname 대체)
// ─────────────────────────────────────────────
const __filename = fileURLToPath(import.meta.url);
const __dirname  = dirname(__filename);

const ROOT     = resolve(__dirname, '..');
const PKG_PATH = join(ROOT, 'package.json');
const GAS_DIR  = join(ROOT, 'gas');
const API_TS   = join(ROOT, 'src', 'api.ts');

// ─────────────────────────────────────────────
// 1. package.json에서 버전 읽기
// ─────────────────────────────────────────────
const pkg     = JSON.parse(readFileSync(PKG_PATH, 'utf8'));
const version = pkg.version;        // e.g. "0.8.6"
const vTag    = `v${version}`;      // e.g. "v0.8.6"

if (!version || !/^\d+\.\d+/.test(version)) {
  console.error(`[bump-version] package.json version 형식 오류: "${version}"`);
  process.exit(1);
}

console.log(`[bump-version] 버전 적용: ${vTag}`);

let totalReplaced = 0;

// ─────────────────────────────────────────────
// 2. gas/*.gs — 상단 주석 버전 치환
//    대상 패턴: "ED's MVP - <설명> vX.X.X..."
//    예: " * ED's MVP - KIS Price Refresh v0.8.5.1"
// ─────────────────────────────────────────────
// g 플래그: 파일 내 여러 줄에서 모두 매칭
const GS_VER_RE = /(ED['']s MVP\s*-[^\n\r]*?)\s+v\d+[\d.]*/g;

const gsFiles = readdirSync(GAS_DIR).filter(f => f.endsWith('.gs'));

for (const filename of gsFiles) {
  const filepath = join(GAS_DIR, filename);
  const original = readFileSync(filepath, 'utf8');

  let count = 0;
  const updated = original.replace(GS_VER_RE, (_, prefix) => {
    count++;
    return `${prefix} ${vTag}`;
  });

  if (count > 0 && updated !== original) {
    writeFileSync(filepath, updated, 'utf8');
    console.log(`  [GAS] ${filename}: ${count}곳 → ${vTag}`);
    totalReplaced += count;
  }
}

// ─────────────────────────────────────────────
// 3. src/api.ts — APP_VERSION 상수 치환
//    대상: export const APP_VERSION = '0.x.x';
//          export const APP_VERSION = "0.x.x";
// ─────────────────────────────────────────────
if (existsSync(API_TS)) {
  const original = readFileSync(API_TS, 'utf8');
  const API_VER_RE = /(export\s+const\s+APP_VERSION\s*=\s*)(['"])[^'"]+\2/;

  let count = 0;
  const updated = original.replace(API_VER_RE, (_, prefix, q) => {
    count++;
    return `${prefix}${q}${version}${q}`;
  });

  if (count > 0 && updated !== original) {
    writeFileSync(API_TS, updated, 'utf8');
    console.log(`  [TS]  api.ts: APP_VERSION → '${version}'`);
    totalReplaced += count;
  } else if (count === 0) {
    console.warn(`  [TS]  api.ts: APP_VERSION 패턴 미발견 — 수동 확인 필요`);
  }
}

// ─────────────────────────────────────────────
// 4. 완료
// ─────────────────────────────────────────────
console.log(`[bump-version] 완료: 총 ${totalReplaced}곳 치환`);
