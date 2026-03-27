const fs = require('fs');
const content = fs.readFileSync('src/App.tsx', 'utf8');
let stack = [];
for (let i = 0; i < content.length; i++) {
  const char = content[i];
  if (char === '(' || char === '{' || char === '[') {
    stack.push({ char, line: content.substring(0, i).split('\n').length });
  } else if (char === ')' || char === '}' || char === ']') {
    if (stack.length === 0) {
      console.log('Extra closing bracket at line', content.substring(0, i).split('\n').length);
      break;
    }
    const last = stack.pop();
    if ((char === ')' && last.char !== '(') ||
        (char === '}' && last.char !== '{') ||
        (char === ']' && last.char !== '[')) {
      console.log('Mismatch at line', content.substring(0, i).split('\n').length, 'expected', last.char, 'got', char);
      break;
    }
  }
}
if (stack.length > 0) {
  console.log('Unclosed brackets:', stack);
}
