#!/usr/bin/env python3
"""restore_si_fixes_v2.py - Restores SI fixes with new marker."""
import sys, os, shutil, datetime, ast

TARGET = sys.argv[1] if len(sys.argv) > 1 else "migrate.py"
if not os.path.exists(TARGET): print(f"X Cannot find {TARGET}"); sys.exit(1)

with open(TARGET, encoding='utf-8') as f: src = f.read()

MARKER = "# SI_FIXES_RESTORED_V2"
if MARKER in src: print("Already applied."); sys.exit(0)

backup = f"{TARGET}.bak_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
shutil.copy2(TARGET, backup)

JS = """\
<script>
/* SI: Enter=linebreak + remove duplicate boxes (restored v2) */
(function(){
  function removeDups(){
    document.querySelectorAll('.si-tl-content').forEach(function(c){
      var b=c.querySelectorAll('.si-tl-body');
      for(var i=1;i<b.length;i++)b[i].remove();
    });
  }
  function siEnter(){
    var el=document.getElementById('content');
    if(!el||el._siEnterV3)return;
    el._siEnterV3=true;
    el.addEventListener('keydown',function(e){
      if(e.key!=='Enter'||e.shiftKey)return;
      var t=e.target;
      if(!t||!t.closest||!t.closest('.si-card'))return;
      if(t.classList&&t.classList.contains('si-action-text'))return;
      e.preventDefault();e.stopPropagation();
      document.execCommand('insertLineBreak');
    },true);
  }
  if(document.readyState!=='loading'){removeDups();siEnter();}
  else document.addEventListener('DOMContentLoaded',function(){removeDups();siEnter();});
})();
</script>
"""

ns=src.find('NOTES_HTML="""')
hc=src.find('</html>',ns)
src=src[:hc]+JS+src[hc:]
src=MARKER+"\n"+src

try: ast.parse(src); print("OK")
except SyntaxError as e: print(f"X {e}"); sys.exit(1)

with open(TARGET,'w',encoding='utf-8') as f: f.write(src)
print("Done. Restart + hard refresh.")