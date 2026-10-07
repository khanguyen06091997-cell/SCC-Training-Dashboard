import base64,json
b64=lambda p:'data:image/png;base64,'+base64.b64encode(open(p,'rb').read()).decode()
h=open('head.html').read(); b=open('body.html').read().replace('__LOGO_W__',b64('logo_w.png')).replace('__LOGO_N__',b64('logo_n.png'))
a=open('app.js').read(); smp=json.dumps(json.load(open('sample_v9.json')),ensure_ascii=False,separators=(',',':'))
open('scc-training-dashboard.html','w').write(h+b+'<script>window.__SAMPLE__='+smp+';</script>\n<script>'+a+'</script>\n')
