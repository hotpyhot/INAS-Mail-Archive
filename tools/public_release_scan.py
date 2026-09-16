from pathlib import Path
import re, sys
ROOT = Path(__file__).resolve().parents[1]
EXCLUDE_DIRS = {'.git','.venv','build','dist','release','__pycache__'}
TEXT_EXT = {'.py','.md','.txt','.json','.bat','.ps1','.yml','.yaml','.toml','.ini','.cfg','.spec'}
email_re = re.compile(r'(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b')
user_path_re = re.compile(r'(?i)[A-Z]:\\Users\\([^\\\s\"\']+)')
secret_value_re = re.compile(r'''(?ix)\b(api[_-]?key|access[_-]?token|client[_-]?secret|password|private[_-]?key)\b\s*[:=]\s*["']?([^\s"']{8,})''')
allowed_email_domains = {'example.com'}
allowed_usernames = {'Example'}
issues=[]
for p in ROOT.rglob('*'):
    if not p.is_file() or any(x in EXCLUDE_DIRS for x in p.parts) or p.suffix.lower() not in TEXT_EXT:
        continue
    text=p.read_text(encoding='utf-8',errors='ignore')
    for n,line in enumerate(text.splitlines(),1):
        for m in email_re.finditer(line):
            if m.group(0).lower().split('@')[-1] not in allowed_email_domains:
                issues.append((p.relative_to(ROOT),n,'non-example email',m.group(0)))
        for m in user_path_re.finditer(line):
            if m.group(1) not in allowed_usernames:
                issues.append((p.relative_to(ROOT),n,'user-specific path',m.group(0)))
        for m in secret_value_re.finditer(line):
            value=m.group(2)
            if value.lower() not in {'examplevalue','changeme123'}:
                issues.append((p.relative_to(ROOT),n,'possible secret value',m.group(0)))
for item in issues:
    print(f'{item[0]}:{item[1]}: {item[2]}: {item[3]}')
print(f'Potential findings: {len(issues)}')
sys.exit(1 if issues else 0)
