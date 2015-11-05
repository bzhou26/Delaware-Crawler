import spidermonkey

import requests
import PyV8
js = spidermonkey.runtime()
js_ctx = js.new_context()
script = requests.urlopen('http://etherhack.co.uk/hashing/whirlpool/js/whirlpool.js').read()
js_ctx.eval_script(script)
js_ctx.eval_script('var s = "abc"')
js_ctx.eval_script('print(HexWhirlpool(s))')
