#!/usr/bin/env python3
"""Builds the Snack & Drink Inventory app from the SP Inventory app.
Same engine (Excel spreadsheet as the database, backups, Dropbox phone version),
with its own name, look, storage, categories (Drinks / Snacks) and typing rules.
After SP Inventory changes:  python3 tools/make_from_sp.py /path/to/sp-repo"""
import re, sys, os, hashlib, base64
OUT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NAME = 'Snack & Drink Inventory'
PHONE_URL = 'https://rumerlaw.github.io/snacks/'
DBX_KEY = 'vyunkynze20qjde'  # the Snack & Drink Dropbox app (its own app, separate from SP Inventory's)

def must(s, a, b, n=1):
    c = s.count(a)
    if c != n: sys.exit('make_snacks: expected %d of %r, found %d' % (n, a[:90], c))
    return s.replace(a, b)

def transform(html, phone):
    head_end = html.index('<body>')
    # ---------- split: markup vs scripts ----------
    parts = re.split(r'(<script>.*?</script>)', html, flags=re.S)
    out = []
    for p in parts:
        if p.startswith('<script>'): out.append(('js', p))
        else: out.append(('html', p))
    s_js = lambda f: [(k, f(v) if k == 'js' else v) for k, v in out]
    s_html = lambda f: [(k, f(v) if k == 'html' else v) for k, v in out]
    joined = lambda o: ''.join(v for k, v in o)

    # ---- markup: name, wording, colours ----
    def markup(h):
        h = h.replace('<title>SP Inventory</title>', '<title>Snack &amp; Drink Inventory</title>')
        h = h.replace('SP Inventory', 'Snack &amp; Drink Inventory')
        h = h.replace('Loading your plants…', 'Loading your stock…')
        h = h.replace('placeholder="Golden Barrel, 6 inch pot, candle…"', 'placeholder="Coke, Doritos, Gatorade…"')
        h = h.replace('>Genus / type<input id="a-genus" type="text" placeholder="Filled in for plants">', '>Size / flavor<input id="a-genus" type="text" placeholder="Optional, e.g. 20 oz or Cool Ranch">')
        h = h.replace('aria-label="Order the plants"', 'aria-label="Order the products"')
        h = h.replace('<option value="plants">Plants first</option>', '<option value="plants">Drinks first</option>')
        h = h.replace('>Products first</option>', '>Snacks first</option>')
        h = h.replace('<option value="genus">Genus A–Z</option>', '<option value="genus">Size / flavor A–Z</option>')
        h = h.replace('placeholder="Joe\'s Nursery', 'placeholder="Costco, Sam\'s Club')
        h = h.replace('aria-label="How to show the plants"', 'aria-label="How to show the products"')
        h = h.replace('(Cacti, Pots and Summary sheets)', '(Stock, Summary and Shortcuts sheets)')
        h = h.replace('<meta name="theme-color" content="#050806">', '<meta name="theme-color" content="#050815">')
        return h
    out = s_html(markup)

    # ---- colours: SP green -> snack blue (CSS sits in the markup part) ----
    COL = {'#050806':'#050815','#0C1410':'#0C1328','#E3F5EA':'#E3ECFA','#7FA38E':'#8296C2','#4A7A62':'#4A63A0','#3DDC84':'#4C9AFF',
           '#03100A':'#020B1F','#0F2A1B':'#0F1C3A','#060A08':'#060914','#0E1612':'#0E1527','#121D18':'#121B33','#EEF7F1':'#EEF3FC',
           '#A3BFAF':'#A3B4D6','#0F2419':'#0F1A33','#2ECC71':'#3B82F6','#0C120E':'#0C1222','#0A110D':'#0A0F1E','#070B08':'#070A14',
           '#D9DDD4':'#D6DCE8','#E4E7DF':'#E2E7F1','#E7EAE2':'#E6EAF3','#EAECE5':'#E9EDF5','#1B2620':'#18213A','#4E5D53':'#4D5B78',
           '#1F8A55':'#2456C9','#E2EEE5':'#E1E8F6','#D5D9D0':'#D2D9E6','#CDD2C8':'#CAD2E2','#7D8C82':'#7C88A3'}
    RGBA = {'150,205,175':'150,175,235','31,138,85':'36,86,201','61,220,132':'76,154,255','40,70,55':'40,55,95','30,50,40':'30,40,70'}
    def colours(h):
        h = re.sub(r'#[0-9A-Fa-f]{6}\b', lambda m: COL.get(m.group(0).upper(), m.group(0)), h)
        for a, b in RGBA.items(): h = h.replace('rgba(' + a, 'rgba(' + b)
        return h
    out = s_html(colours)

    js = joined([(k, v) for k, v in out])  # work on the whole page from here, carefully
    # ---- storage: own names, so the two apps never share sign-ins, files or settings (same website) ----
    js = re.sub(r"'(?:sp|cactus)-([a-z0-9-]*[:']?)", lambda m: "'snk-" + m.group(1), js)
    js = js.replace("indexedDB.open('snk-xl'", "indexedDB.open('snk-xl'")
    js = js.replace('SP Inventory', NAME)

    # ---- categories: Drinks or Snacks; no plant genus guessing ----
    js = must(js, "  function guessCategory(name,genus){return 'Plants'}",
      "  /* Drinks or Snacks, from the name; a category you name yourself always wins */\n"
      "  var SNACK_RE=/\\b(chocolate|choc|candy|candies|gum|mints?|chips?|crisps|doritos|fritos|cheetos|lays|ruffles|pringles|takis|cheez|cheez-it|crackers?|cookies?|bars?|granola|popcorn|pretzels?|nuts|peanuts|almonds|trail|jerky|gummy|gummies|haribo|skittles|m&ms|snickers|twix|kitkat|kit kat|reeses|hershey|oreos?|fig|krispies|treats?|muffins?|donuts?|pastry|pastries|sandwich|burrito|noodles|ramen|cup of noodles|fruit snacks|bananas?|apples?|oranges?|fruit|gardettos|sunchips|goldfish|ice breakers|extra)\\b/i;\n"
      "  var DRINK_RE=/\\b(coke|cola|sprite|fanta|pepsi|water|waters|soda|sodas|pop|gatorade|powerade|juice|juices|tea|teas|iced tea|coffee|cold brew|latte|lemonade|monster|red bull|redbull|drink|drinks|pellegrino|san pellegrino|perrier|topo chico|sparkling|seltzer|la croix|lacroix|bubly|dr pepper|mountain dew|mtn dew|energy|celsius|rockstar|bang|prime|body ?armor|vitamin ?water|smartwater|arrowhead|dasani|aquafina|fiji|evian|essentia|propel|core|liquid death|milk|chocolate milk|kombucha|ginger ale|root beer|7 ?up|a&w|sunkist|squirt|crush|snapple|arizona|gold peak|pure leaf|v8|yerba|olipop|poppi|zevia|electrolit|pedialyte|shake|smoothie|beer|wine)\\b/i;\n"
      "  function guessCategory(name,genus){var t=String(name||'');if(DRINK_RE.test(t)&&!/\\b(chocolate|candy|bar|bars|chips|gummy|gummies)\\b/i.test(t.replace(/chocolate milk/ig,'')))return 'Drinks';if(SNACK_RE.test(t))return 'Snacks';return DRINK_RE.test(t)?'Drinks':'Snacks'}")
    js = must(js, "  function guessGenus(name){", "  function guessGenus(name){return '';")
    js = must(js, "  function genusOfWord(w){", "  function genusOfWord(w){return '';")
    # sort: "Drinks first" / "Snacks first"
    js = must(js, "var pa=catKey(a.category)==='plant'?0:1,pb=catKey(b.category)==='plant'?0:1", "var pa=catKey(a.category)==='drink'?0:1,pb=catKey(b.category)==='drink'?0:1")
    # no plant one-time fixes (every product into Plants; size word moved first)
    js = must(js, "if(canWrite&&!localStorage.getItem(k2)&&", "if(false&&canWrite&&!localStorage.getItem(k2)&&")
    js = must(js, "if(canWrite&&!localStorage.getItem(k3)&&", "if(false&&canWrite&&!localStorage.getItem(k3)&&")
    # shortcuts: sizes only (no plant ones like H = Haworthia)
    m = re.search(r"var SHORT_BUILTIN=NP\(\{.*?\}\)", js, re.S)
    if not m: sys.exit('make_snacks: shortcuts not found')
    js = js[:m.start()] + "var SHORT_BUILTIN=NP({sm:'Small',sml:'Small',med:'Medium',md:'Medium',lg:'Large',lrg:'Large',xl:'Extra Large',xs:'Extra Small',pk:'Pack',pks:'Packs'})" + js[m.end():]
    # typo fixer: only fix 1-letter slips, and never brand words ("sprite" is not "price")
    js = must(js, "var max=l.length<=5?1:2,best=null,bd=99;", "var max=1,best=null,bd=99;if(l.length<5)return w;")
    js = must(js, "if(l.length<4||plant[l]||VOCAB.indexOf(l)>=0||DONTFIX.indexOf(' '+l+' ')>=0)return w;",
              "if(l.length<4||plant[l]||VOCAB.indexOf(l)>=0||DONTFIX.indexOf(' '+l+' ')>=0||DRINK_RE.test(l)||SNACK_RE.test(l))return w;")
    # Spreadsheet column: Size / flavor
    js = js.replace("'Genus / type'", "'Size / flavor'")
    # ---- your snack commands: "+24 coke", "new Sprite, .55, 24, 2", "price coke 2.50", "lost 1 banana" ----
    js = must(js, "  function listLine(line){\n",
      "  function listLine(line){\n"
      "    var sn=snackLine(line);if(sn)return sn;\n")
    js = must(js, "  function listLines(t){",
      "  /* the Snack & Drink ways of typing, turned into the normal commands */\n"
      "  function snackLine(line){\n"
      "    var t=String(line).trim().replace(/(^|[^\\d])\\.(\\d)/g,'$10.$2');\n"
      "    var a=t.match(/^\\+\\s*(\\d+)\\s+(.*[a-z].*)$/i);if(a)return ['bought '+a[1]+' '+expandShort(a[2]).trim()];\n"
      "    var n=t.match(/^(?:new|add)\\s+([^,]+?)\\s*,\\s*\\$?(\\d+(?:\\.\\d+)?)\\s*,\\s*(\\d+)\\s*,\\s*\\$?(\\d+(?:\\.\\d+)?)\\s*(?:,\\s*(drinks?|snacks?))?\\s*$/i);\n"
      "    if(n)return ['\\u0001'+JSON.stringify({name:expandShort(n[1]).trim(),cost:parseFloat(n[2]),qty:parseInt(n[3],10),price:parseFloat(n[4]),cat:n[5]?(/^d/i.test(n[5])?'Drinks':'Snacks'):null})];\n"
      "    var p=t.match(/^(?:set\\s+)?price\\s+(.+?)\\s+(?:to\\s+)?\\$?(\\d+(?:\\.\\d+)?)$/i);if(p)return ['set '+expandShort(p[1]).trim()+' to '+p[2]];\n"
      "    return null;\n"
      "  }\n"
      "  function listLines(t){")
    js = must(js, "  async function interpret(text,spoken){\n",
      "  async function interpret(text,spoken){\n"
      "    /* \"new Name, cost, how many, price\": the name is used exactly as typed */\n"
      "    if(/^\\u0001/.test(String(text))){var d=JSON.parse(String(text).slice(1)),ex=allItems.filter(function(i){return norm(i.name)===norm(d.name)})[0];\n"
      "      if(ex)return {kind:'restock',item:ex,qty:d.qty,cost:d.cost,price:d.price,genus:null,vendor:null,text:'Add '+d.qty+' '+fullName(ex)+'. Stock goes from '+ex.qty+' to '+(ex.qty+d.qty)+'.'};\n"
      "      var nm=title(d.name);return {kind:'new',costTotal:null,name:nm,genus:'',category:d.cat||guessCategory(nm),vendor:'',qty:d.qty,cost:d.cost,price:d.price,text:'New: '+nm+', '+d.qty+' in stock, cost '+money(d.cost)+' each, sells '+money(d.price)+' each.'}}\n"
      "    /* \"lost 1 banana\", \"threw out 2 chips\": fewer in stock, not a sale */\n"
      "    var lm=String(text).trim().match(/^(?:lost|waste|wasted|trash|trashed|toss|tossed|threw out|thrown out|expired|broke|broken|spoiled|damaged)\\s+(\\d+)\\s+(.+)$/i);\n"
      "    if(lm){var ln=parseInt(lm[1],10),li=find(expandShort(lm[2]).trim());if(!li)return {error:'I could not find a product called “'+lm[2]+'”.'};\n"
      "      if(!(ln>0))return {error:'Use a number above 0.'};if(ln>li.qty)return {error:'Only '+li.qty+' '+fullName(li)+' in stock.'};\n"
      "      return {kind:'set',item:li,field:'qty',value:li.qty-ln,text:'Remove '+ln+' '+fullName(li)+' (lost or thrown out, not sold). Stock goes from '+li.qty+' to '+(li.qty-ln)+'.'}}\n")
    # ---- wording ----
    W = [("a plant called", "a product called"), ("which plant you sold. Say its name.", "which product you sold. Say its name."),
         ("Which plant did you get? Say its name.", "Which product did you get? Say its name."),
         ("No plants match “", "Nothing matches “"), (">Delete plant</button>", ">Delete</button>"),
         ("Try “sold 3 golden barrel”, “set bunny ears price to 10”, or “add 20 bunny ears, cost 3, sell 9”.", "Try “sold 2 gatorade”, “set coke to 2.50”, or “bought 24 coke at 0.55 each sell for 2 each”."),
         ("Try “bought 6 golden barrel at 5 sell for 12”.", "Try “bought 24 coke at 0.55 each sell for 2 each”."),
         ("for example “bought 6 golden barrel at 6 from Desert Greens sell for 15” or “bought 6 candles at 3 sell for 10”", "for example “bought 35 coke for 19.29 sell for 2” (the pack price) or “24 gatorade 0.56 2”"),
         ("Example plants added. Try saying “sold 3 golden barrel”.", "Example products added. Try typing “sold 3 coke”."),
         ("A cactus seller typed or said: ", "A snack and drink seller typed or said: "),
         ("You turn one sentence from a cactus seller into JSON.", "You turn one sentence from a snack and drink seller into JSON."),
         ("Find the plant (or pot) being added.", "Find the product being added."), ("If no plant name is in the sentence", "If no product name is in the sentence"),
         ("'Plants already in the list: ", "'Products already in the list: ")]
    for a, b in W: js = js.replace(a, b)
    # example products
    js = must(js, "var ex=[['Golden Barrel','Echinocactus',6,15,5],['Old Man Cactus','Cephalocereus',4.5,12,15],['Bunny Ears','Opuntia',3,9,25]];",
              "var ex=[['Coke','',0.55,2,35],['Gatorade','',0.56,2,28],['Cheez-It','',0.31,1,45]];")
    js = must(js, "snoozedQty:null,category:'Plants',example:true", "snoozedQty:null,category:guessCategory(x[0]),example:true")
    # Latin-name splitting is for plants; snack names stay as typed
    js = must(js, "  function splitName(v){\n    v=String(v||'').trim();\n", "  function splitName(v){\n    v=String(v||'').trim();return {name:v,sci:''};\n")
    # typing and speaking examples
    m = re.search(r"eg:\['Add cactus called bg.*?'\]", js)
    if not m: sys.exit('make_snacks: speak examples not found')
    js = js[:m.start()] + "eg:['I sold 2 gatorade','Got 24 coke for 19 dollars, selling them for 2 dollars each','Set coke price to 2.50','Lost 1 banana']" + js[m.end():]
    js = must(js, "placeholder:'sold 3 golden barrel',", "placeholder:'sold 2 gatorade',")
    m = re.search(r"help:'Quickest way to add:.*?',\n", js)
    if not m: sys.exit('make_snacks: type help not found')
    js = js[:m.start()] + "help:'Quickest way to add: how many, the name, what you paid each, then the sell price each. For example “24 gatorade 0.56 2” adds 24 Gatorade bought for $0.56 each, selling for $2. Paid for the whole pack? Type “bought 35 coke for 19.29 sell for 2” and it works out $0.55 each. To sell, type “sold 2 gatorade” or “-2 gatorade”. “+24 coke” adds stock and “lost 1 banana” takes one out without counting a sale. Put one on each line to add many at once.',\n" + js[m.end():]
    m = re.search(r"eg:\['1 trich 5 20'.*?'\]", js)
    if not m: sys.exit('make_snacks: type examples not found')
    js = js[:m.start()] + "eg:['24 gatorade 0.56 2','sold 2 coke','-1 cheez-it','+24 coke','bought 35 coke for 19.29 sell for 2','new Sprite, .55, 24, 2','price coke 2.50','lost 1 banana']" + js[m.end():]
    # "Coke Zero" is a name, not the number 0
    js = must(js, "var NV=NP({zero:0,", "var NV=NP({")
    # "Gardetto's", not "Gardetto'S"
    js = must(js, "function title(s){return String(s).trim().replace(/\\s+/g,' ').replace(/\\b([a-z])/g,function(m){return m.toUpperCase()})",
                  "function title(s){return String(s).trim().replace(/\\s+/g,' ').replace(/(^|[^A-Za-z'\\u2019])([a-z])/g,function(m,a,b){return a+b.toUpperCase()})")
    js = must(js, "  function fixTypos(text){\n", "  function fixTypos(text){\n    if(/(^|\\n)\\s*(?:new|add)\\s+[^,\\n]+,/i.test(String(text)))return String(text);\n")
    # ---- cheat sheet ----
    m = re.search(r"  function cheatHtml\(\)\{.*?\n  \}\n", js, re.S)
    if not m: sys.exit('make_snacks: cheat sheet not found')
    cheat = r"""  function cheatHtml(){
    function ex(t){return '<code>'+h(t)+'</code>'}
    var base=[['SM','Small'],['MED','Medium'],['LG','Large'],['XL','Extra Large'],['PK','Pack']];
    var mine=Object.keys(USERSHORT).sort().map(function(k){return [k.toUpperCase(),USERSHORT[k]]});
    var sc=base.concat(mine).map(function(x){return '<li><b>'+h(x[0])+'</b> '+h(x[1])+'</li>'}).join('');
    var op=true;try{op=localStorage.getItem('snk-cheat')!=='closed'}catch(e){}
    return '<details class="cheatwrap" id="cheatd"'+(op?' open':'')+'><summary>How to type it · cheat sheet</summary><div class="cheat">'+
      '<div class="cs"><h4>Add, the quickest way</h4>'+ex('24 gatorade 0.56 2')+'<ul><li>how many → <b>24</b></li><li>name → <b>gatorade</b></li><li>paid each → <b>$0.56</b></li><li>sell price each → <b>$2</b></li></ul>'+
        '<h4>Add from the pack price</h4>'+ex('bought 35 coke for 19.29 sell for 2')+'<ul><li><b>$19.29</b> is for the whole pack, so each one cost <b>$0.55</b></li></ul>'+ex('new Sprite, .55, 24, 2')+'<ul><li>name, cost each, how many, sell price</li></ul>'+
        '<h4>Add many at once</h4><ul><li>one per line: press <b>Shift + Enter</b> for a new line</li><li>or put <b>;</b> between them: '+ex('24 coke 0.55 2; 12 doritos 0.37 1')+'</li></ul></div>'+
      '<div class="cs"><h4>Sell</h4>'+ex('sold 2 gatorade')+'<ul><li>or quicker: '+ex('-2 gatorade')+'</li><li>sold for a different price? add it: '+ex('-2 coke 1.50')+'</li></ul>'+
        '<h4>More stock</h4>'+ex('+24 coke')+'<ul><li>adds 24 to what you have</li></ul>'+
        '<h4>Lost or thrown out</h4>'+ex('lost 1 banana')+'<ul><li>takes it out of stock, not counted as a sale</li></ul>'+
        '<h4>Change</h4>'+ex('set coke to 2.50')+'<ul><li>or '+ex('price coke 2.50')+' changes the sell price</li><li>cost: '+ex('coke cost 0.55')+'</li><li>stock: '+ex('coke stock 30')+'</li></ul>'+
        '<h4>Delete</h4>'+ex('delete coke')+'<ul><li>deletes right away</li><li>press <b>Undo</b> to bring it back</li></ul></div>'+
      '<div class="cs"><h4>Shortcuts</h4><ul class="sc">'+sc+'</ul><ul><li>add your own: type '+ex('cz means coke zero')+' in the box and press <b>Enter</b></li><li>remove one: '+ex('delete shortcut cz')+'</li><li>see them all: '+ex('shortcuts')+'</li></ul></div>'+
      '</div><p class="fine" style="margin-top:6px">Press <b>Enter</b> to save. Typos are fine.</p></details>';
  }
"""
    js = js[:m.start()] + cheat + js[m.end():]

    if phone:
        js = must(js, '"redirect": "https://rumerlaw.github.io/sp/"', '"redirect": "' + PHONE_URL + '"')
        js = must(js, '"key": "5dqv7rvruwnnu5f"', '"key": "' + DBX_KEY + '"')
        js = js.replace("prompt('Name for the new spreadsheet','" + NAME + "')", "prompt('Name for the new spreadsheet','Snack and Drink Inventory')")
    # ---- security policy: fingerprints of the new scripts ----
    hs = ["'sha256-" + base64.b64encode(hashlib.sha256(x.encode('utf-8')).digest()).decode() + "'" for x in re.findall(r'<script>(.*?)</script>', js, re.S)]
    js, k = re.subn(r"script-src 'self'(?: 'sha256-[A-Za-z0-9+/=]+')+", "script-src 'self' " + ' '.join(hs), js)
    if k != 1: sys.exit('make_snacks: security policy not found')
    return js

if __name__ == '__main__':
    SP = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(OUT), 'sp')
    desk = transform(open(os.path.join(SP, 'index.html'), encoding='utf-8').read(), False)
    open(os.path.join(OUT, 'index.html'), 'w', encoding='utf-8').write(desk)
    ph = transform(open(os.path.join(SP, 'docs', 'index.html'), encoding='utf-8').read(), True)
    open(os.path.join(OUT, 'docs', 'index.html'), 'w', encoding='utf-8').write(ph)
    print('made index.html and docs/index.html from', SP)
