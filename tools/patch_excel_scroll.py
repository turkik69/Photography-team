from pathlib import Path

p=Path('index.html')
s=p.read_text(encoding='utf-8')

needle='function ensureXLSXLoaded()'
start=s.find(needle)
if start<0:
    raise SystemExit('ensureXLSXLoaded not found')
brace=s.find('{',start)
if brace<0:
    raise SystemExit('ensureXLSXLoaded opening brace not found')
depth=0
end=None
in_s=in_d=in_t=False
esc=False
for i in range(brace,len(s)):
    ch=s[i]
    if esc:
        esc=False
        continue
    if ch=='\\' and (in_s or in_d or in_t):
        esc=True
        continue
    if in_s:
        if ch=="'": in_s=False
        continue
    if in_d:
        if ch=='"': in_d=False
        continue
    if in_t:
        if ch=='`': in_t=False
        continue
    if ch=="'": in_s=True; continue
    if ch=='"': in_d=True; continue
    if ch=='`': in_t=True; continue
    if ch=='{': depth+=1
    elif ch=='}':
        depth-=1
        if depth==0:
            end=i+1
            break
if end is None:
    raise SystemExit('ensureXLSXLoaded closing brace not found')

new="""function ensureXLSXLoaded(){
  if(window.XLSX) return Promise.resolve(window.XLSX);
  return new Promise((resolve,reject)=>{
    const script=document.createElement('script');
    script.src='https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
    script.onload=()=>window.XLSX?resolve(window.XLSX):reject(new Error('XLSX style library unavailable'));
    script.onerror=()=>reject(new Error('تعذر تحميل مكتبة Excel'));
    document.head.appendChild(script);
  });
}"""
s=s[:start]+new+s[end:]
p.write_text(s,encoding='utf-8')
