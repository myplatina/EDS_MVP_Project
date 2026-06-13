import fs from 'fs';
import path from 'path';

// 1. Single Source of Truth (package.json) 버전 읽기
const packageJsonPath = path.resolve(process.cwd(), 'package.json');
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
const currentVersion = packageJson.version; // 무조건 "0.8.7" 확보

console.log(`\n🚀 [Version Sync] package.json 기준 동기화 타겟: v${currentVersion}`);

// 2. 루트 및 gas 폴더 전수 조사
const targetDirs = [process.cwd(), path.resolve(process.cwd(), 'gas')];

targetDirs.forEach(dir => {
  if (!fs.existsSync(dir)) return;

  const files = fs.readdirSync(dir).filter(f => f.endsWith('.gs'));
  files.forEach(file => {
    const filePath = path.join(dir, file);
    let content = fs.readFileSync(filePath, 'utf8');
    let isModified = false;

    // 규칙 A: 파일 최상단 주석 주위의 "v0.x.x" 패턴 강제 포획 및 치환 [cite: 1]
    const headerRegex = /(ED's MVP - .*? v)\d+\.\d+\.\d+(\.\d+)?/gi;
    if (headerRegex.test(content)) {
      content = content.replace(headerRegex, `$1${currentVersion}`);
      isModified = true;
    }

    // 규칙 B: 내부 제어 상수 "version: 'x.x.x'" 패턴 완벽 치환 (0.8.5.2 기만 코드 전원 사살) 
    const constantRegex = /(version\s*:\s*['"])\d+\.\d+\.\d+(\.\d+)?(['"])/gi;
    if (constantRegex.test(content)) {
      content = content.replace(constantRegex, `$1${currentVersion}$3`);
      isModified = true;
    }

    if (isModified) {
      fs.writeFileSync(filePath, content, 'utf8');
      console.log(`✨ 싱크 완료 -> ${file} (v${currentVersion})`);
    }
  });
});
console.log(`🏁 [Sync Complete] 모든 백엔드 원장의 버전 스탬프가 v${currentVersion} 로 대통합되었습니다.\n`);