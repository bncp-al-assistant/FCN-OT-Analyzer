// samples/*.xlsx → js/sample-data.js (브라우저에서 파일 서버 없이도 예시를 열 수 있도록 base64로 포함)
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..', 'samples');
const files = fs.readdirSync(dir).filter(f => /\.xlsx$/i.test(f)).sort();
const data = files.map(name => ({ name, b64: fs.readFileSync(path.join(dir, name)).toString('base64') }));
const out = '/* 자동 생성 파일: npm run embed */\nvar SAMPLE_FILES = ' + JSON.stringify(data) + ';\n';
fs.writeFileSync(path.join(__dirname, '..', 'js', 'sample-data.js'), out);
console.log('embedded', files.length, 'sample file(s)');
