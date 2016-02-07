<%'20160207130204-/fwx/fwx.asp%>
<%Response.Buffer=True
Function RegularExpression()
Set RegularExpression=New RegExp
RegularExpression.Global=True
RegularExpression.IgnoreCase=True
End Function
Dim r
Function StartRegularExpressions()
Set r=RegularExpression()
Set StartRegularExpressions=r
End Function
Call StartRegularExpressions()
Const rxFiles="[a-zA-Z0-9\.\,\@\-\_\#\~]+"
Const rxNotFiles="[^a-zA-Z0-9\.\,\@\-\_\#\~]+"
Const rxNumbers="[0-9]+"
Const rxNotNumbers="[^0-9]+"
Const rxNumeric="[0-9\-\.]+"
Const rxNotNumeric="[^0-9\-\.]+"
Const rxSymbols="\W+"
Const rxNotSymbols="\w+"
Const rxLinks="(?<HTML><a[^>]*href\s*=\s*[\""\']?(?<HRef>[^""'>\s]*)[\""\']?[^>]*>(?<Title>[^<]+|.*?)?</a\s*>)"
Const rxComments="<!--((?!-->)[\s\S])*?-->"
Const rxProperties="\w+\ *\=\ *((\""((?!\"").)*\"")|(\'((?!\').)*\'))"
Const rxEmailAddress="[a-z0-9][a-z0-9\_\.\-\+]{0,}[a-z0-9]@[a-z0-9][a-z0-9_\.-]{0,}[a-z0-9][\.][a-z0-9]{2,4}"
Dim rxEmail:rxEmail="^"&rxEmailAddress&"$"
Dim rxUsername:rxUsername="[a-zA-Z0-9_\.]{4,16}"
Dim rxPassword:rxPassword=".{5,16}"
Dim rxProtocol:rxProtocol="http|https|ftp"
Dim rxDomainTLD:rxDomainTLD="(\.(aero|biz|com|coop|edu|gov|info|int|mil|museum|name|net|org|ac|ad|ae|af|ag|ai|al|am|an|ao|aq|ar|as|at|au|aw|az|ba|bb|bd|be|bf|bg|bh|bi|bj|bm|bn|bo|br|bs|bt|bv|bw|by|bz|ca|cc|cd|cf|cg|ch|ci|ck|cl|cm|cn|co|cr|cs|cu|cv|cx|cy|cz|de|dj|dk|dm|do|dz|ec|ee|eg|eh|er|es|et|fi|fj|fk|fm|fo|fr|ga|gb|gd|ge|gf|gg|gh|gi|gl|gm|gn|gp|gq|gr|gs|gt|gu|gw|gy|hk|hm|hn|hr|ht|hu|id|ie|il|im|in|io|iq|ir|is|it|je|jm|jo|jp|ke|kg|kh|ki|km|kn|kp|kr|kw|ky|kz|la|lb|lc|li|lk|lr|ls|lt|lu|lv|ly|ma|mc|md|mg|mh|mk|ml|mm|mn|mo|mp|mq|mr|ms|mt|mu|mv|mw|mx|my|mz|na|nc|ne|nf|ng|ni|nl|no|np|nr|nu|nz|om|pa|pe|pf|pg|ph|pk|pl|pm|pn|pr|ps|pt|pw|py|qa|re|ro|ru|rw|sa|sb|sc|sd|se|sg|sh|si|sj|sk|sl|sm|sn|so|sr|st|su|sv|sy|sz|tc|td|tf|tg|th|tj|tk|tm|tn|to|tp|tr|tt|tv|tw|tz|ua|ug|uk|um|us|uy|uz|va|vc|ve|vg|vi|vn|vu|wf|ws|ye|yt|yu|za|zm|zr|zw)){1,2}"
Dim rxDomain:rxDomain="([a-zA-Z0-9]"&"[a-zA-Z0-9\-]+"&"\.)*?([a-zA-Z0-9]"&"[a-zA-Z0-9\-]+"&")"&rxDomainTLD&""
Dim rxURL:rxURL="(("&rxProtocol&")\:\/\/)?"&rxDomain&"(:[\d]{1,5})?(/[\w\-\_\.]*)*((\?\w+=\w+)?(&\w+=\w+)*)?"
Function GetDomain(ByVal s)
If Len(s)>0 Then GetDomain=rxSelect(s,rxDomain)
End Function
Function GetTLD(ByVal s)
s=GetDomain(s)
If Len(s)>0 Then GetTLD=rxReplace(rxSelect(s,rxDomainTLD&"$"),"^\.","")
End Function
Function GetTLDLess(ByVal s)
s=GetDomain(s)
If Len(s)>0 Then GetTLDLess=rxSelect(Left(s,Len(s)-Len("."&GetTLD(s))),"[^\.]+$")
End Function
Function GetSubDomain(ByVal s)
s=GetDomain(s)
If Len(s)>0 Then GetSubDomain=rxReplace(s,"^([^\.]+)\..+$","$1")
End Function
Const rxExtDocs="doc|pdf"
Const rxExtHtml="htm|html|asp"
Const rxExtFeeds="xml|rss|rdf|atom|txt|csv|tab|tsv|js|json|xls|vcf|src"
Const rxExtKnown="xml|rss|rdf|atom|txt|csv|tab|tsv|js|json|xls|vcf|src|doc|pdf|htm|html|asp"
Const rxUKPostcode="^(GIR 0AA)|(((A[BL]|B[ABDHLNRSTX]?|C[ABFHMORTVW]|D[ADEGHLNTY]|E[HNX]?|F[KY]|G[LUY]?|H[ADGPRSUX]|I[GMPV]|JE|K[ATWY]|L[ADELNSU]?|M[EKL]?|N[EGNPRW]?|O[LX]|P[AEHLOR]|R[GHM]|S[AEGKLMNOPRSTY]?|T[ADFNQRSW]|UB|W[ADFNRSV]|YO|ZE)[1-9]?[0-9]|((E|N|NW|SE|SW|W)1|EC[1-4]|WC[12])[A-HJKMNPR-Y]|(SW|W)([2-9]|[1-9][0-9])|EC[1-9][0-9]) [0-9][ABD-HJLNP-UW-Z]{2})$"
Const rxCAPostcode="^[ABCEGHJKLMNPRSTVXY]{1}\d{1}[A-Z]{1} *\d{1}[A-Z]{1}\d{1}$"
Const rxUSCounty="^\w{2}$"
Const rxUSPostcode="^(\d{5}-\d{4}|\d{5}|\d{9})$|^([a-zA-Z]\d[a-zA-Z] \d[a-zA-Z]\d)$"
Const rxNIPostcode="^(BT1|BT10|BT11|BT12|BT13|BT14|BT15|BT16|BT17|BT2|BT27|BT28|BT29|BT3|BT38|BT39|BT4|BT40|BT41|BT42|BT43|BT44|BT5|BT53|BT54|BT56|BT57|BT58|BT6|BT7|BT8|BT9|BT60|BT61|BT62|BT63|BT64|BT65|BT66|BT67|BT18|BT19|BT20|BT21|BT22|BT23|BT24|BT25|BT26|BT30|BT31|BT32|BT33|BT34|BT35|BT74|BT92|BT93|BT94|BT45|BT46|BT47|BT49|BT51|BT55|BT68|BT69|BT70|BT71|BT74|BT75|BT76|BT77|BT78|BT79|BT80|BT81|BT82)\s?.*$"
Dim rxStopWords:rxStopWords=""
rxStopWords=rxStopWords&"|a|able|about|above|according|accordingly|across|actually|after|afterwards|again|against|ain't|all|allow|allows|almost|alone|along|already|also|although|always|am|among|amongst|an|and|another|any|anybody|anyhow|anyone|anything|anyway|anyways|anywhere|apart|appear|appreciate|appropriate|approximately|are|area|areas|aren't|around|as|a's|aside|ask|asked|asking|asks|associated|at|available|away|awfully"
rxStopWords=rxStopWords&"|b|back|backed|backing|backs|be|became|because|become|becomes|becoming|been|before|beforehand|began|behind|being|beings|believe|below|beside|besides|best|better|between|beyond|big|both|brief|but|by|"
rxStopWords=rxStopWords&"|c|came|can|cannot|cant|can't|case|cases|cause|causes|certain|certainly|changes|clear|clearly|c'mon|co|com||come|comes|concerning|consequently|consider|considering|contain|containing|contains|corresponding|could|couldn't|course|c's|currently|"
rxStopWords=rxStopWords&"|d|de||definitely|described|despite|did|didn't|differ|different|differently|do|does|doesn't|doing|done|don't|down|downed|downing|downs|downwards|due|during|"
rxStopWords=rxStopWords&"|e|each|early|edu|eg|eight|either|else|elsewhere|en|end|ended|ending|ends|enough|entirely|especially|et|etc|even|evenly|ever|every|everybody|everyone|everything|everywhere|ex|exactly|example|except|"
rxStopWords=rxStopWords&"|f|face|faces|fact|facts|far|felt|few|fifth|find|finds|first|five|followed|following|follows|for|former|formerly|forth|found|four|from|full|fully|further|furthered|furthering|furthermore|furthers|"
rxStopWords=rxStopWords&"|g|gave|general|generally|get|gets|getting|give|given|gives|giving|go|goes|going|gone|good|goods|got|gotten|great|greater|greatest|greetings|group|grouped|grouping|groups|"
rxStopWords=rxStopWords&"|h|had|hadn't|happens|hardly|has|hasn't|have|haven't|having|he|hello|help|hence|her|here|hereafter|hereby|herein|here's|hereupon|hers|herself|he's|hi|high|higher|highest|him|himself|his|hither|hopefully|how|howbeit|however|"
rxStopWords=rxStopWords&"|i|i'd|ie|if|ignored|i'll|i'm|immediate|important|in|inasmuch|inc|indeed|indicate|indicated|indicates|inner|insofar|instead|interest|interested|interesting|interests|into|inward|is|isn't|it|it'd|it'll|its|it's|itself|i've|"
rxStopWords=rxStopWords&"|j|just|"
rxStopWords=rxStopWords&"|k|keep|keeps|kept|kg|kind|km|knew|know|known|knows|"
rxStopWords=rxStopWords&"|l|la|large|largely|last|lately|later|latest|latter|latterly|least|less|lest|let|lets|let's|like|liked|likely|little|long|longer|longest|look|looking|looks|ltd|"
rxStopWords=rxStopWords&"|m|made|mainly|make|making|man|many|may|maybe|me|mean|meanwhile|member|members|men|merely|might|min|ml|mm|more|moreover|most|mostly|mr|mrs|much|must|my|myself|"
rxStopWords=rxStopWords&"|n|name|namely|nd|near|nearly|necessary|need|needed|needing|needs|neither|never|nevertheless|new|newer|newest|next|nine|no|nobody|non|none|noone|nor|normally|not|nothing|novel|now|nowhere|number|numbers|"
rxStopWords=rxStopWords&"|o|obtain|obtained|obviously|of|off|often|oh|ok|okay|old|older|oldest|on|once|one|ones|only|onto|open|opened|opening|opens|or|order|ordered|ordering|orders|other|others|otherwise|ought|our|ours|ourselves|out|outside|over|overall|own|"
rxStopWords=rxStopWords&"|p|part|parted|particular|particularly|parting|parts|per|perhaps|place|placed|places|please|plus|point|pointed|pointing|points|possible|present|presented|presenting|presents|presumably|previously|probably|problem|problems|provides|put|puts|"
rxStopWords=rxStopWords&"|q|que|quite|qv|"
rxStopWords=rxStopWords&"|r|rather|rd|re|really|reasonably|regarding|regardless|regards|relatively|respectively|resulted|resulting|right|room|rooms|"
rxStopWords=rxStopWords&"|s|said|same|saw|say|saying|says|second|secondly|seconds|see|seeing|seem|seemed|seeming|seems|seen|sees|self|selves|sensible|sent|serious|seriously|seven|several|shall|she|should|shouldn't|show|showed|showing|shown|shows|side|sides|significant|significantly|since|six|small|smaller|smallest|so|some|somebody|somehow|someone|something|sometime|sometimes|somewhat|somewhere|soon|sorry|specified|specify|specifying|state|states|still|sub|such|suggest|sup|sure|"
rxStopWords=rxStopWords&"|t|take|taken|tell|tends|th|than|thank|thanks|thanx|that|thats|that's|the|their|theirs|them|themselves|then|thence|there|thereafter|thereby|therefore|therein|theres|there's|thereupon|these|they|they'd|they'll|they're|they've|thing|things|think|thinks|third|this|thorough|thoroughly|those|though|thought|thoughts|three|through|throughout|thru|thus|to|today|together|too|took|toward|towards|tried|tries|truly|try|trying|t's|turn|turned|turning|turns|twice|two|"
rxStopWords=rxStopWords&"|u|un|und|under|unfortunately|unless|unlikely|until|unto|up|upon|us|use|used|useful|uses|using|usually|"
rxStopWords=rxStopWords&"|v|value|various|very|via|viz|vs|"
rxStopWords=rxStopWords&"|w|want|wanted|wanting|wants|was|wasn't|way|ways|we|we'd|welcome|well|we'll|wells|went|were|we're|weren't|we've|what|whatever|what's|when|whence|whenever|where|whereafter|whereas|whereby|wherein|where's|whereupon|wherever|whether|which|while|whither|who|whoever|whole|whom|who's|whose|why|will|willing|wish|with|within|without|wonder|won't|work|worked|working|works|would|wouldn't|www|"
rxStopWords=rxStopWords&"|x|y|year|years|yes|yet|you|you'd|you'll|young|younger|youngest|your|you're|yours|yourself|yourselves|you've|"
rxStopWords=rxStopWords&"|z|zero|"
Dim rxBadWords:rxBadWords=""
rxBadWords=rxBadWords&"|ahole|anus|ash0le|ash0les|asholes|ass|Ass Monkey|Assface|assh0le|assh0lez|asshole|assholes|assholz|asswipe|azzhole|bassterds|bastard|bastards|bastardz|basterds|basterdz|Biatch|bitch|bitches|Blow Job|boffing|butthole|buttwipe|c0ck|c0cks|c0k|Carpet Muncher|cawk|cawks|Clit|cnts|cntz|cock|cockbiter|cockbiters|cock-biter|cock-biters|cockhead|cock-head|cocks|CockSucker|cock-sucker|crap|cum|cunt|cunts|cuntz|damn|darn|dick|dild0|dild0s|dildo|dildos|dilld0|dilld0s|dominatricks|dominatrics|dominatrix|dyke"
rxBadWords=rxBadWords&"|enema|f u c k|f u c k e r|fag|fag1t|faget|fagg1t|faggit|faggot|fagit|fags|fagz|faig|faigs|fart|flipping the bird|fuck|fucker|fuckin|fucking|fucks|Fudge Packer|fuk|Fukah|Fuken|fuker|Fukin|Fukk|Fukkah|Fukken|Fukker|Fukkin|g00k|gay|gayboy|gaygirl|gays|gayz|God-damned|h00r|h0ar|h0re|hells|hoar|hoor|hoore|jackoff|jap|japs|jerk-off|jisim|jiss|jizm|jizz|knob|knobs|knobz|kunt|kunts|kuntz|Lesbian|Lezzian|Lipshits|Lipshitz|masochist|masokist|massterbait|masstrbait|masstrbate|masterbaiter|masterbate|masterbates"
rxBadWords=rxBadWords&"|Motha Fucker|Motha Fuker|Motha Fukkah|Motha Fukker|Mother Fucker|Mother Fukah|Mother Fuker|Mother Fukkah|Mother Fukker|mother-fucker|Mutha Fucker|Mutha Fukah|Mutha Fuker|Mutha Fukkah|Mutha Fukker|n1gr|nastt|nigger;|nigur;|niiger;|niigr;|orafis|orgasim;|orgasm|orgasum|oriface|orifice|orifiss|packi|packie|packy|paki|pakie|paky|pecker|peeenus|peeenusss|peenus|peinus|pen1s|penas|penis|penis-breath|penus|penuus|Phuc|Phuck|Phuk|Phuker|Phukker|polac|polack|polak|Poonani|pr1c|pr1ck|pr1k|pusse|pussee|pussy|puuke"
rxBadWords=rxBadWords&"|puuker|queer|queers|queerz|qweers|qweerz|qweir|recktum|rectum|retard|sadist|scank|schlong|screwing|semen|sex|sexy|Sh!t|sh1t|sh1ter|sh1ts|sh1tter|sh1tz|shit|shits|shitter|Shitty|Shity|shitz|Shyt|Shyte|Shytty|Shyty|skanck|skank|skankee|skankey|skanks|Skanky|slut|sluts|Slutty|slutz|son-of-a-bitch|tit|turd|va1jina|vag1na|vagiina|vagina|vaj1na|vajina|vullva|vulva|w0p|wh00r|wh0re|whore|xrated|"
Function rxTest(ByVal s,ByVal p)
If s="" Then
rxTest=False
Else
Err.Clear():On Error Resume Next
r.Pattern=p
rxTest=r.Test(TOE(s))
If Err.Number<>0 Then
If IsTest Then DBG.FatalError "rxTest ERROR p='"&p&"' s='"&s&"'"
Err.Clear()
rxTest=False
End If
End If
End Function
Function rxEscape(ByVal s)
If s<>"" Then
r.Pattern="(\W)"
s=r.Replace(s,"\$1")
End If
rxEscape=s
End Function
Function rxMatches(ByVal s,ByVal p)
r.Pattern=p
Set rxMatches=r.Execute(s)
End Function
Function rxReplace(ByVal s,ByVal p,ByVal s2)
If s<>"" Then
r.Pattern=p
rxReplace=r.Replace(s,s2)
End If
End Function
Function rxReplaceAll(ByVal s,ByVal p,ByVal s2)
If s<>"" Then
r.Pattern=p
Dim i:i=1
Do While r.Test(s)And i<=5
i=i+1
s=r.Replace(s,s2)
Loop
End If
rxReplaceAll=s
End Function
Function rxFind(ByVal s,ByVal p)
On Error Resume Next
Dim output:output=""
If s<>"" Then
r.Pattern=p
If r.Test(s)Then
Dim aM:Set aM=r.Execute(s)
If IsObject(aM)Then
output=aM(0)
End If
End If
End If
rxFind=output
End Function
Function rxSelect(ByVal s,ByVal p):rxSelect=rxFind(s,p):End Function
Function RemoveSymbols(ByVal s)
If s<>"" Then
s=rxReplace(s,rxSymbols,"")
End If
RemoveSymbols=s
End Function
Function NumbersOnly(ByVal str)
If s<>"" Then
s=rxReplace(s,rxNotNumbers,"")
End If
NumbersOnly=s
End Function
Function Dictionary()
Set Dictionary=Server.CreateObject("Scripting.Dictionary")
End Function
Function FSObject()
Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
End Function
Function ConnectionObject()
Set ConnectionObject=Server.CreateObject("ADODB.Connection")
End Function
Function Recordset()
Dim objRecordSet
CreateRS objRecordSet
Set Recordset=objRecordSet
End Function
Function XMLDOM()
Set XMLDOM=Server.CreateObject("MSXML2.DOMDocument.6.0")
End Function
Function XMLHTTP()
Set XMLHTTP=Server.CreateObject("MSXML2.XMLHTTP.6.0")
End Function%><script language="Javascript" runat="server">
Object.prototype.get||(Object.prototype.get=function(t){return this[t]}),Object.prototype.set||(Object.prototype.set=function(t,e){if("unknown"==typeof e)try{e=new VBArray(e).toArray()}catch(r){return}this[t]=e}),Object.prototype.purge||(Object.prototype.purge=function(t){delete this[t]}),Object.prototype.keys||(Object.prototype.keys=function(){var t=new ActiveXObject("Scripting.Dictionary");for(var e in this)this.hasOwnProperty(e)&&t.add(e,this[e]);return t.keys()}),Object.prototype.values||(Object.prototype.values=function(){var t=new ActiveXObject("Scripting.Dictionary");for(var e in this)this.hasOwnProperty(e)&&t.add(e,this[e]);return t.items()}),String.prototype.sanitize||(String.prototype.sanitize=function(t,e){var r=t.length,u=this;if(r!==e.length)throw new TypeError("Invalid procedure call. Both arrays should have the same size.");for(var n=0;r>n;n++){var o=new RegExp(t[n],"g");u=u.replace(o,e[n])}return u}),String.prototype.substitute||(String.prototype.substitute=function(t,e){return this.replace(e||/\\?\{([^{}]+)\}/g,function(e,r){return"\\"==e.charAt(0)?e.slice(1):void 0!=t[r]?t[r]:""})}),"object"!=typeof JSON&&(JSON={}),function(){"use strict";function f(t){return 10>t?"0"+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=0,rx_escapable.test(t)?'"'+t.replace(rx_escapable,function(t){var e=meta[t];return"string"==typeof e?e:"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})+'"':'"'+t+'"'}function str(t,e){var r,u,n,o,i,a=gap,s=e[t];switch(s&&"object"==typeof s&&"function"==typeof s.toJSON&&(s=s.toJSON(t)),"function"==typeof rep&&(s=rep.call(e,t,s)),typeof s){case"string":return quote(s);case"number":return isFinite(s)?String(s):"null";case"boolean":case"null":return String(s);case"object":if(!s)return"null";if(gap+=indent,i=[],"[object Array]"===Object.prototype.toString.apply(s)){for(o=s.length,r=0;o>r;r+=1)i[r]=str(r,s)||"null";return n=0===i.length?"[]":gap?"[\n"+gap+i.join(",\n"+gap)+"\n"+a+"]":"["+i.join(",")+"]",gap=a,n}if(rep&&"object"==typeof rep)for(o=rep.length,r=0;o>r;r+=1)"string"==typeof rep[r]&&(u=rep[r],n=str(u,s),n&&i.push(quote(u)+(gap?": ":":")+n));else for(u in s)Object.prototype.hasOwnProperty.call(s,u)&&(n=str(u,s),n&&i.push(quote(u)+(gap?": ":":")+n));return n=0===i.length?"{}":gap?"{\n"+gap+i.join(",\n"+gap)+"\n"+a+"}":"{"+i.join(",")+"}",gap=a,n}}var rx_one=/^[\],:{}\s]*$/,rx_two=/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,rx_four=/(?:^|:|,)(?:\s*\[)+/g,rx_escapable=/[\\\"\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;"function"!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+"-"+f(this.getUTCMonth()+1)+"-"+f(this.getUTCDate())+"T"+f(this.getUTCHours())+":"+f(this.getUTCMinutes())+":"+f(this.getUTCSeconds())+"Z":null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value);var gap,indent,meta,rep;"function"!=typeof JSON.stringify&&(meta={"\b":"\\b"," ":"\\t","\n":"\\n","\f":"\\f","\r":"\\r",'"':'\\"',"\\":"\\\\"},JSON.stringify=function(t,e,r){var u;if(gap="",indent="","number"==typeof r)for(u=0;r>u;u+=1)indent+=" ";else"string"==typeof r&&(indent=r);if(rep=e,e&&"function"!=typeof e&&("object"!=typeof e||"number"!=typeof e.length))throw new Error("JSON.stringify");return str("",{"":t})}),"function"!=typeof JSON.parse&&(JSON.parse=function(text,reviver){function walk(t,e){var r,u,n=t[e];if(n&&"object"==typeof n)for(r in n)Object.prototype.hasOwnProperty.call(n,r)&&(u=walk(n,r),void 0!==u?n[r]=u:delete n[r]);return reviver.call(t,e,n)}var j;if(text=String(text),rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})),rx_one.test(text.replace(rx_two,"@").replace(rx_three,"]").replace(rx_four,"")))return j=eval("("+text+")"),"function"==typeof reviver?walk({"":j},""):j;throw new SyntaxError("JSON.parse")})}(),function(){function t(t){return/^[_:A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD][-.·_:0-9A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD\u0300-\u036F\u203F-\u2040]*$/.test(t)?t:"_INVALID_NAMESTARTCHAR_"+t}function e(t){return t.sanitize(["&","<",">","'",'"'],["&","&lt;","&gt;","&apos;","&quot;"])}function r(u,n){var o,i,a,s=[],p=[];switch(typeof u){case"object":if(null===u)s.push("<{tag}/>".substitute({tag:t(n)}));else if(u.length)if(p=u,0===p.length)s.push("<{tag}/>".substitute({tag:t(n)}));else for(i=0,a=p.length;a>i;i++)s.push(r(p[i],n));else{s.push("<{tag}".substitute({tag:t(n)})),p=[];for(o in u)u.hasOwnProperty(o)&&("@"===o.charAt(0)?s.push(" {param}='{content}'".substitute({param:t(o.substr(1)),content:e(u[o].toString())})):p.push(o));if(0===p.length)s.push("/>");else{for(s.push(">"),i=0,a=p.length;a>i;i++)o=p[i],s.push("#text"===o?e(u[o]):"#cdata"===o?"<![CDATA[{code}]]>".substitute({code:u[o]}):r(u[o],o));s.push("</{tag}>".substitute({tag:t(n)}))}}break;default:var f=String(u);s.push(0===f.length?"<{tag}/>".substitute({tag:t(n)}):"<{tag}>{value}</{tag}>".substitute({tag:t(n),value:e(f)}))}return s.join("")}"function"!=typeof JSON.toXML&&(JSON.toXML=function(t,e){var u=[];e&&u.push("<{tag}>".substitute({tag:e}));for(var n in t)t.hasOwnProperty(n)&&u.push(r(t[n],n));return e&&u.push("</{tag}>".substitute({tag:e})),u.join("")})}(),function(){"function"!=typeof JSON.minify&&(JSON.minify=function(t){var e,r,u,n,o=/"|(\/\*)|(\*\/)|(\/\/)|\n|\r/g,i=!1,a=!1,s=!1,p=[],f=0,c=0;for(o.lastIndex=0;e=o.exec(t);)u=RegExp.leftContext,n=RegExp.rightContext,a||s||(r=u.substring(c),i||(r=r.replace(/(\n|\r|\s)*/g,"")),p[f++]=r),c=o.lastIndex,'"'!=e[0]||a||s?"/*"!=e[0]||i||a||s?"*/"!=e[0]||i||!a||s?"//"!=e[0]||i||a||s?"\n"!=e[0]&&"\r"!=e[0]||i||a||!s?a||s||/\n|\r|\s/.test(e[0])||(p[f++]=e[0]):s=!1:s=!0:a=!1:a=!0:(r=u.match(/(\\)*$/),i&&r&&r[0].length%2!=0||(i=!i),c--,n=t.substring(c));return p[f++]=n,p.join("")})}();
</script><%Function ParseJson(ByVal strJson)
Dim json
Set json=New VbsJson
json.Decode(strJson)
Set ParseJson=json
End Function
Function GenerateJson(ByVal obj)
Dim json,o
Set json=New VbsJson
o=json.Encode(obj)
GenerateJson=o
Set json=Nothing
End Function
Class VbsJson
Private Whitespace,NumberRegex,StringChunk
Private b,f,r,n,t
Private Sub Class_Initialize
Whitespace=" "&vbTab&vbCr&vbLf
b=ChrW(8)
f=vbFormFeed
r=vbCr
n=vbLf
t=vbTab
Set NumberRegex=New RegExp
NumberRegex.Pattern="(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
NumberRegex.Global=False
NumberRegex.MultiLine=True
NumberRegex.IgnoreCase=True
Set StringChunk=New RegExp
StringChunk.Pattern="([\s\S]*?)([""\\\x00-\x1f])"
StringChunk.Global=False
StringChunk.MultiLine=True
StringChunk.IgnoreCase=True
End Sub
Public Function Encode(ByRef obj)
Dim buf,i,c,g
Set buf=CreateObject("Scripting.Dictionary")
Select Case VarType(obj)
Case vbNull
buf.Add buf.Count,"null"
Case vbBoolean
If obj Then
buf.Add buf.Count,"true"
Else
buf.Add buf.Count,"false"
End If
Case vbInteger,vbLong,vbSingle,vbDouble
buf.Add buf.Count,obj
Case vbString
buf.Add buf.Count,""""
For i=1 To Len(obj)
c=Mid(obj,i,1)
Select Case c
Case """" buf.Add buf.Count,"\"""
Case "\" buf.Add buf.Count,"\\"
Case "/" buf.Add buf.Count,"/"
Case b buf.Add buf.Count,"\b"
Case f buf.Add buf.Count,"\f"
Case r buf.Add buf.Count,"\r"
Case n buf.Add buf.Count,"\n"
Case t buf.Add buf.Count,"\t"
Case Else
If AscW(c)>=0 And AscW(c)<=31 Then
c=Right("0"& Hex(AscW(c)),2)
buf.Add buf.Count,"\u00"&c
Else
buf.Add buf.Count,c
End If
End Select
Next
buf.Add buf.Count,""""
Case vbArray+vbVariant
g=True
buf.Add buf.Count,"["
For Each i In obj
If g Then g=False Else buf.Add buf.Count,","
buf.Add buf.Count,Encode(i)
Next
buf.Add buf.Count,"]"
Case vbObject
If TypeName(obj)="Dictionary" Then
g=True
buf.Add buf.Count,"{"
For Each i In obj
If g Then g=False Else buf.Add buf.Count,","
buf.Add buf.Count,""""&i&""":"&Encode(obj(i))
Next
buf.Add buf.Count,"}"
Else
Err.Raise 8732,,"None dictionary object"
End If
Case Else
buf.Add buf.Count,""""&CStr(obj)&""""
End Select
Encode=Join(buf.Items,"")
End Function
Public Function Decode(ByRef str)
Dim idx
idx=SkipWhitespace(str,1)
If Mid(str,idx,1)="{" Then
Set Decode=ScanOnce(str,1)
Else
Decode=ScanOnce(str,1)
End If
End Function
Private Function ScanOnce(ByRef str,ByRef idx)
Dim c,ms
idx=SkipWhitespace(str,idx)
c=Mid(str,idx,1)
If c="{" Then
idx=idx+1
Set ScanOnce=ParseObject(str,idx)
Exit Function
ElseIf c="[" Then
idx=idx+1
ScanOnce=ParseArray(str,idx)
Exit Function
ElseIf c="""" Then
idx=idx+1
ScanOnce=ParseString(str,idx)
Exit Function
ElseIf c="n" And StrComp("null",Mid(str,idx,4))=0 Then
idx=idx+4
ScanOnce=Null
Exit Function
ElseIf c="t" And StrComp("true",Mid(str,idx,4))=0 Then
idx=idx+4
ScanOnce=True
Exit Function
ElseIf c="f" And StrComp("false",Mid(str,idx,5))=0 Then
idx=idx+5
ScanOnce=False
Exit Function
End If
Set ms=NumberRegex.Execute(Mid(str,idx))
If ms.Count=1 Then
idx=idx+ms(0).Length
ScanOnce=CDbl(ms(0))
Exit Function
End If
Err.Raise 8732,,"No JSON object could be ScanOnced"
End Function
Private Function ParseObject(ByRef str,ByRef idx)
Dim c,key,value
Set ParseObject=CreateObject("Scripting.Dictionary")
idx=SkipWhitespace(str,idx)
c=Mid(str,idx,1)
If c="}" Then
Exit Function
ElseIf c<>"""" Then
Err.Raise 8732,,"Expecting property name"
End If
idx=idx+1
Do
key=ParseString(str,idx)
idx=SkipWhitespace(str,idx)
If Mid(str,idx,1)<>":" Then
Err.Raise 8732,,"Expecting : delimiter"
End If
idx=SkipWhitespace(str,idx+1)
If Mid(str,idx,1)="{" Then
Set value=ScanOnce(str,idx)
Else
value=ScanOnce(str,idx)
End If
ParseObject.Add key,value
idx=SkipWhitespace(str,idx)
c=Mid(str,idx,1)
If c="}" Then
Exit Do
ElseIf c<>"," Then
Err.Raise 8732,,"Expecting , delimiter"
End If
idx=SkipWhitespace(str,idx+1)
c=Mid(str,idx,1)
If c<>"""" Then
Err.Raise 8732,,"Expecting property name"
End If
idx=idx+1
Loop
idx=idx+1
End Function
Private Function ParseArray(ByRef str,ByRef idx)
Dim c,values,value
Set values=CreateObject("Scripting.Dictionary")
idx=SkipWhitespace(str,idx)
c=Mid(str,idx,1)
If c="]" Then
ParseArray=values.Items
Exit Function
End If
Do
idx=SkipWhitespace(str,idx)
If Mid(str,idx,1)="{" Then
Set value=ScanOnce(str,idx)
Else
value=ScanOnce(str,idx)
End If
values.Add values.Count,value
idx=SkipWhitespace(str,idx)
c=Mid(str,idx,1)
If c="]" Then
Exit Do
ElseIf c<>"," Then
Err.Raise 8732,,"Expecting , delimiter"
End If
idx=idx+1
Loop
idx=idx+1
ParseArray=values.Items
End Function
Private Function ParseString(ByRef str,ByRef idx)
Dim chunks,content,terminator,ms,esc,char
Set chunks=CreateObject("Scripting.Dictionary")
Do
Set ms=StringChunk.Execute(Mid(str,idx))
If ms.Count=0 Then
Err.Raise 8732,,"Unterminated string starting"
End If
content=ms(0).Submatches(0)
terminator=ms(0).Submatches(1)
If Len(content)>0 Then
chunks.Add chunks.Count,content
End If
idx=idx+ms(0).Length
If terminator="""" Then
Exit Do
ElseIf terminator<>"\" Then
Err.Raise 8732,,"Invalid control character"
End If
esc=Mid(str,idx,1)
If esc<>"u" Then
Select Case esc
Case """" char=""""
Case "\" char="\"
Case "/" char="/"
Case "b" char=b
Case "f" char=f
Case "n" char=n
Case "r" char=r
Case "t" char=t
Case Else Err.Raise 8732,,"Invalid escape"
End Select
idx=idx+1
Else
char=ChrW("&H"&Mid(str,idx+1,4))
idx=idx+5
End If
chunks.Add chunks.Count,char
Loop
ParseString=Join(chunks.Items,"")
End Function
Private Function SkipWhitespace(ByRef str,ByVal idx)
Do While idx<=Len(str)And InStr(Whitespace,Mid(str,idx,1))>0
idx=idx+1
Loop
SkipWhitespace=idx
End Function
End Class
Function STR2JSON(ByVal s)
If s="" Then Exit Function
Dim x
Set x=New VbsJson
Set STR2JSON=x.Decode(s)
Set x=Nothing
End Function
Function GetJSON(ByVal f):Set GetJSON=STR2JSON(GetResponse(f)):End Function
Function ReadJSON(ByVal f):Set ReadJSON=STR2JSON(ReadFile(f)):End Function
Function STR2Array(ByVal s)
If s="" Then Exit Function
Dim x
Set x=New VbsJson
STR2Array=x.Decode(s)
Set x=Nothing
End Function
Function GetArray(ByVal f):GetArray=STR2Array(GetResponse(f)):End Function
Function ReadArray(ByVal f):ReadArray=STR2Array(ReadFile(f)):End Function
class ClsCache
private prefix
private property get cachebase
set cachebase=nothing
if not isArray(application(prefix&name))then
application(prefix&name)=array(server.createObject("scripting.dictionary"))
end if
set cachebase=application(prefix&name)(0)
end property
public name
public interval
public intervalValue
public maxSlots
public maxItemSize
public sub class_initialize()
name=""
interval="h"
intervalValue=1
maxSlots=10
maxItemSize=100000
prefix="GL_Cache_"
end sub
public sub store(identifier,item)
if name="" then lib.error("ICE 4598: Name is needed for caching.")
if not isObject(item)and not isArray(item)then
if len(item)>maxItemSize then exit sub
end if
set c=cachebase
expires=dateadd(interval,intervalValue,now())
packedItem=array(expires,item)
if c.exists(identifier)then c.remove(identifier)
c.add identifier,packedItem
reOrganize()
end sub
public sub clear()
set c=cachebase
c.removeAll()
end sub
public function getItem(identifier)
set c=cachebase
if(c.exists(identifier))then
cachedItem=c(identifier)
if not itemExpired(cachedItem)then
getItem=cachedItem(1)
else
c.remove(identifier)
end if
end if
end function
private sub reOrganize()
set c=cachebase
if c.count>maxSlots then
removedAtLeastOne=false
for each identifier in c.keys
rf=c(identifier)
if itemExpired(rf)then
c.remove(identifier)
removedAtLeastOne=true
end if
next
if not removedAtLeastOne then
identifiers=c.keys
c.remove(identifiers(0))
end if
end if
end sub
private function itemExpired(cachedItemArray)
itemExpired=(cachedItemArray(0)<now())
end function
end class
Class fwxCache
Public Prefix
Public Expiry
Public Method
Private Sub Class_Initialize()
Prefix="@@CACHE"
Expiry=(60*24)
Method=1
End Sub
Public Sub Clear()
If Method=2 Then
DeleteFolder "/"&Prefix&"/"
Else
For Each s In Application.Contents
If Left(s,Len(Prefix))=(Prefix)Then
Application.Contents.Remove(s)
End If
Next
End If
End Sub
Public Function Key(ByVal s)
s=rxReplace(s,rxFiles,"-")
Key=Prefix&s&""
End Function
Public Function File(ByVal s)
s=rxReplace(s,rxFiles,"-")
File="/"&Prefix&"/"&s&".txt"
End Function
Public Property Let Item(ByVal sID,ByVal sValue)
Application.Contents(Key(sID)&"_D")=Now()
If Method=2 Then
WriteToFile File(sID)
Else
Application.Contents(Key(sID))=sValue
End If
End Property
Public Default Property Get Item(ByVal sID)
Dim output:output=""
Dim dDate,iHowLong
iHowLong=-1
dDate=Application.Contents(Key(sID)&"_D")
If dDate<>"" Then
dDate=CDate(dDate)
iHowLong=DateDiff("n",dDate,Now())
End If
If iHowLong<0 Or iHowLong>=Expiry Then
If Method=2 Then
DeleteFile File(sID)
Else
Application.Contents.Remove(Key(sID))
End If
Else
If Method=2 Then
output=ReadFile(File(sID))
Else
output=Application.Contents(Key(sID))
End If
End If
Item=output
End Property
End Class
Dim cache:Set cache=New fwxCache
Class fwxSOAP
Public Req,Res
Public doc
Public Base
Public SRV
Public USR
Public PWD
Public Function Q(ByVal d)
Q=Me.Send("GET",d)
End Function
Public Function F(ByVal d)
F=Me.Send("POST",d)
End Function
Public Function Send(ByVal m,ByVal d)
Me.Req=d
Dim x:Set x=Server.CreateObject("Microsoft.XMLHTTP")
x.Open m,Me.SRV,False,Me.USR,Me.PWD
x.setRequestHeader "Content-Type","application/xml"
x.send d
If Err.Description<>"" Or Err.Number<>0 Then
With Response
.ContentType="text/plain"
.Write "ERROR ("&met&"): "&Err.Description&vbCRLf
.Write Me.Req&""
.End()
End With
End If
Err.Clear()
Me.Res=x.responseText
Set x=Nothing
Set doc=LoadXml(Me.Res)
Send=Me.Res
End Function
Public Sub Purge()
With Response
.Clear()
.ContentType="application/xml"
.Write Me.xml
.End()
End With
End Sub
Public Default Property Get Data(ByVal sNode)
On Error Resume Next
If Left(sNode,2)="//" Then
Data=doc.selectSingleNode(sNode&"").Text&""
Else
Data=doc.selectSingleNode(Base&sNode&"").Text&""
End If
End Property
Public Property Let Data(ByVal sNode,ByVal sValue)
On Error Resume Next
If Left(sNode,2)="//" Then
doc.selectSingleNode(sNode&"").Text=sValue
Else
doc.selectSingleNode(Base&sNode&"").Text=sValue
End If
If Err.Description<>"" Or Err.Number<>0 Then
If Base<>"" Then
Me.Insert sNode,sValue
End If
End If
End Property
Public Function InsertIn(ByRef no,ByVal sNode,ByVal sValue)
Set nn=doc.createElement(sNode&"")
nn.Text=sValue&""
no.appendChild nn
Set InsertIn=nn
End Function
Public Function Insert(ByVal sNode,ByVal sValue)
Dim no,nn
If Base<>"" Then
Set no=doc.selectSingleNode(Left(Base,Len(Base)-1)&"")
Else
Set no=doc
End If
Set Insert=InsertIn(no,sNode,sValue)
End Function
Function ListNodes(ByVal sNode)
Dim aNodes,sN,so
aNodes=Split(sNode,",")
For Each sN In aNodes
If so<>"" Then so=so&", "
so=so&sN&": "&Me.Data(sN)&""
Next
ListNodes=so
End Function
Public Property Get NodeText(ByRef oNode,ByVal sNode)
On Error Resume Next
NodeText=oNode.selectSingleNode(sNode&"").Text&""
End Property
Public Property Let NodeText(ByRef oNode,ByVal sNode,ByVal sValue)
On Error Resume Next
oNode.selectSingleNode(sNode&"").Text=sValue
End Property
Public Property Get Nodes(ByVal sNode)
Set Nodes=doc.selectNodes(Base&sNode&"")
End Property
Function Xml()
Xml=doc.xml&""
End Function
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
End Sub
End Class
Class DynamicArray
Private aData
Private Sub Class_Initialize()
Redim aData(0)
End Sub
Public Default Property Get Item(iPos)
If iPos<LBound(aData)or iPos>UBound(aData)then
Exit Property
End If
Item=aData(iPos)
End Property
Public Property Get Data()
Data=aData
End Property
Public Property Let Item(iPos,varValue)
If iPos<LBound(aData)Then Exit Property
If iPos>UBound(aData)then
Redim Preserve aData(iPos)
aData(iPos)=varValue
Else
aData(iPos)=varValue
End If
End Property
Public Sub Add(ByVal v)
Me.Item(Me.Length)=v
End Sub
Public Sub Clear()
Erase aData
ReDim aData(0)
End Sub
Public Function StartIndex()
StartIndex=LBound(aData)
End Function
Public Function StopIndex()
StopIndex=UBound(aData)
End Function
Public Function Length()
Length=StopIndex+1
End Function
Public Sub Delete(iPos)
If iPos<LBound(aData)or iPos>UBound(aData)then
Exit Sub
End If
Dim iLoop
For iLoop=iPos to UBound(aData)-1
aData(iLoop)=aData(iLoop+1)
Next
Redim Preserve aData(UBound(aData)-1)
End Sub
End Class
Class fwxMD5
Private BITS_TO_A_BYTE
Private BYTES_TO_A_WORD
Private BITS_TO_A_WORD
Private m_lOnBits(30)
Private m_l2Power(30)
Private Sub Class_Initialize()
BITS_TO_A_BYTE=8
BYTES_TO_A_WORD=4
BITS_TO_A_WORD=32
m_lOnBits(0)=CLng(1)
m_lOnBits(1)=CLng(3)
m_lOnBits(2)=CLng(7)
m_lOnBits(3)=CLng(15)
m_lOnBits(4)=CLng(31)
m_lOnBits(5)=CLng(63)
m_lOnBits(6)=CLng(127)
m_lOnBits(7)=CLng(255)
m_lOnBits(8)=CLng(511)
m_lOnBits(9)=CLng(1023)
m_lOnBits(10)=CLng(2047)
m_lOnBits(11)=CLng(4095)
m_lOnBits(12)=CLng(8191)
m_lOnBits(13)=CLng(16383)
m_lOnBits(14)=CLng(32767)
m_lOnBits(15)=CLng(65535)
m_lOnBits(16)=CLng(131071)
m_lOnBits(17)=CLng(262143)
m_lOnBits(18)=CLng(524287)
m_lOnBits(19)=CLng(1048575)
m_lOnBits(20)=CLng(2097151)
m_lOnBits(21)=CLng(4194303)
m_lOnBits(22)=CLng(8388607)
m_lOnBits(23)=CLng(16777215)
m_lOnBits(24)=CLng(33554431)
m_lOnBits(25)=CLng(67108863)
m_lOnBits(26)=CLng(134217727)
m_lOnBits(27)=CLng(268435455)
m_lOnBits(28)=CLng(536870911)
m_lOnBits(29)=CLng(1073741823)
m_lOnBits(30)=CLng(2147483647)
m_l2Power(0)=CLng(1)
m_l2Power(1)=CLng(2)
m_l2Power(2)=CLng(4)
m_l2Power(3)=CLng(8)
m_l2Power(4)=CLng(16)
m_l2Power(5)=CLng(32)
m_l2Power(6)=CLng(64)
m_l2Power(7)=CLng(128)
m_l2Power(8)=CLng(256)
m_l2Power(9)=CLng(512)
m_l2Power(10)=CLng(1024)
m_l2Power(11)=CLng(2048)
m_l2Power(12)=CLng(4096)
m_l2Power(13)=CLng(8192)
m_l2Power(14)=CLng(16384)
m_l2Power(15)=CLng(32768)
m_l2Power(16)=CLng(65536)
m_l2Power(17)=CLng(131072)
m_l2Power(18)=CLng(262144)
m_l2Power(19)=CLng(524288)
m_l2Power(20)=CLng(1048576)
m_l2Power(21)=CLng(2097152)
m_l2Power(22)=CLng(4194304)
m_l2Power(23)=CLng(8388608)
m_l2Power(24)=CLng(16777216)
m_l2Power(25)=CLng(33554432)
m_l2Power(26)=CLng(67108864)
m_l2Power(27)=CLng(134217728)
m_l2Power(28)=CLng(268435456)
m_l2Power(29)=CLng(536870912)
m_l2Power(30)=CLng(1073741824)
End Sub
Private Function LShift(lValue,iShiftBits)
If iShiftBits=0 Then
LShift=lValue
Exit Function
ElseIf iShiftBits=31 Then
If lValue And 1 Then
LShift=&H80000000
Else
LShift=0
End If
Exit Function
ElseIf iShiftBits<0 Or iShiftBits>31 Then
Err.Raise 6
End If
If(lValue And m_l2Power(31-iShiftBits))Then
LShift=((lValue And m_lOnBits(31-(iShiftBits+1)))*m_l2Power(iShiftBits))Or &H80000000
Else
LShift=((lValue And m_lOnBits(31-iShiftBits))*m_l2Power(iShiftBits))
End If
End Function
Private Function RShift(lValue,iShiftBits)
If iShiftBits=0 Then
RShift=lValue
Exit Function
ElseIf iShiftBits=31 Then
If lValue And &H80000000 Then
RShift=1
Else
RShift=0
End If
Exit Function
ElseIf iShiftBits<0 Or iShiftBits>31 Then
Err.Raise 6
End If
RShift=(lValue And &H7FFFFFFE)\m_l2Power(iShiftBits)
If(lValue And &H80000000)Then
RShift=(RShift Or(&H40000000\m_l2Power(iShiftBits-1)))
End If
End Function
Private Function RotateLeft(lValue,iShiftBits)
RotateLeft=LShift(lValue,iShiftBits)Or RShift(lValue,(32-iShiftBits))
End Function
Private Function AddUnsigned(lX,lY)
Dim lX4
Dim lY4
Dim lX8
Dim lY8
Dim lResult
lX8=lX And &H80000000
lY8=lY And &H80000000
lX4=lX And &H40000000
lY4=lY And &H40000000
lResult=(lX And &H3FFFFFFF)+(lY And &H3FFFFFFF)
If lX4 And lY4 Then
lResult=lResult Xor &H80000000 Xor lX8 Xor lY8
ElseIf lX4 Or lY4 Then
If lResult And &H40000000 Then
lResult=lResult Xor &HC0000000 Xor lX8 Xor lY8
Else
lResult=lResult Xor &H40000000 Xor lX8 Xor lY8
End If
Else
lResult=lResult Xor lX8 Xor lY8
End If
AddUnsigned=lResult
End Function
Private Function F(x,y,z)
F=(x And y)Or((Not x)And z)
End Function
Private Function G(x,y,z)
G=(x And z)Or(y And(Not z))
End Function
Private Function H(x,y,z)
H=(x Xor y Xor z)
End Function
Private Function I(x,y,z)
I=(y Xor(x Or(Not z)))
End Function
Private Sub FF(a,b,c,d,x,s,ac)
a=AddUnsigned(a,AddUnsigned(AddUnsigned(F(b,c,d),x),ac))
a=RotateLeft(a,s)
a=AddUnsigned(a,b)
End Sub
Private Sub GG(a,b,c,d,x,s,ac)
a=AddUnsigned(a,AddUnsigned(AddUnsigned(G(b,c,d),x),ac))
a=RotateLeft(a,s)
a=AddUnsigned(a,b)
End Sub
Private Sub HH(a,b,c,d,x,s,ac)
a=AddUnsigned(a,AddUnsigned(AddUnsigned(H(b,c,d),x),ac))
a=RotateLeft(a,s)
a=AddUnsigned(a,b)
End Sub
Private Sub II(a,b,c,d,x,s,ac)
a=AddUnsigned(a,AddUnsigned(AddUnsigned(I(b,c,d),x),ac))
a=RotateLeft(a,s)
a=AddUnsigned(a,b)
End Sub
Private Function ConvertToWordArray(sMessage)
Dim lMessageLength
Dim lNumberOfWords
Dim lWordArray()
Dim lBytePosition
Dim lByteCount
Dim lWordCount
Dim lByteValue
Dim lMessageType
Const MODULUS_BITS=512
Const CONGRUENT_BITS=448
lMessageType=Vartype(sMessage)
Select Case lMessageType
Case 8:lMessageLength=Len(sMessage)
Case 8209:lMessageLength=LenB(sMessage)
Case Else:Err.Raise-1,"MD5","Unknown Type passed to MD5 function ("&lMessageType&","&sMessage&")"
End Select
lNumberOfWords=(((lMessageLength+((MODULUS_BITS-CONGRUENT_BITS)\BITS_TO_A_BYTE))\(MODULUS_BITS\BITS_TO_A_BYTE))+1)*(MODULUS_BITS\BITS_TO_A_WORD)
ReDim lWordArray(lNumberOfWords-1)
lBytePosition=0
lByteCount=0
Do Until lByteCount>=lMessageLength
lWordCount=lByteCount\BYTES_TO_A_WORD
lBytePosition=(lByteCount Mod BYTES_TO_A_WORD)*BITS_TO_A_BYTE
Select Case lMessageType
Case 8:lByteValue=Asc(Mid(sMessage,lByteCount+1,1))
Case 8209:lByteValue=AscB(MidB(sMessage,lByteCount+1,1))
End Select
lWordArray(lWordCount)=lWordArray(lWordCount)Or LShift(lByteValue,lBytePosition)
lByteCount=lByteCount+1
Loop
lWordCount=lByteCount\BYTES_TO_A_WORD
lBytePosition=(lByteCount Mod BYTES_TO_A_WORD)*BITS_TO_A_BYTE
lWordArray(lWordCount)=lWordArray(lWordCount)Or LShift(&H80,lBytePosition)
lWordArray(lNumberOfWords-2)=LShift(lMessageLength,3)
lWordArray(lNumberOfWords-1)=RShift(lMessageLength,29)
ConvertToWordArray=lWordArray
End Function
Private Function WordToHex(lValue)
Dim lByte
Dim lCount
For lCount=0 To 3
lByte=RShift(lValue,lCount*BITS_TO_A_BYTE)And m_lOnBits(BITS_TO_A_BYTE-1)
WordToHex=WordToHex&Right("0"& Hex(lByte),2)
Next
End Function
Public Function MD5(sMessage)
Dim x
Dim k
Dim AA
Dim BB
Dim CC
Dim DD
Dim a
Dim b
Dim c
Dim d
Const S11=7
Const S12=12
Const S13=17
Const S14=22
Const S21=5
Const S22=9
Const S23=14
Const S24=20
Const S31=4
Const S32=11
Const S33=16
Const S34=23
Const S41=6
Const S42=10
Const S43=15
Const S44=21
x=ConvertToWordArray(sMessage)
a=&H67452301
b=&HEFCDAB89
c=&H98BADCFE
d=&H10325476
For k=0 To UBound(x)Step 16
AA=a
BB=b
CC=c
DD=d
FF a,b,c,d,x(k+0),S11,&HD76AA478
FF d,a,b,c,x(k+1),S12,&HE8C7B756
FF c,d,a,b,x(k+2),S13,&H242070DB
FF b,c,d,a,x(k+3),S14,&HC1BDCEEE
FF a,b,c,d,x(k+4),S11,&HF57C0FAF
FF d,a,b,c,x(k+5),S12,&H4787C62A
FF c,d,a,b,x(k+6),S13,&HA8304613
FF b,c,d,a,x(k+7),S14,&HFD469501
FF a,b,c,d,x(k+8),S11,&H698098D8
FF d,a,b,c,x(k+9),S12,&H8B44F7AF
FF c,d,a,b,x(k+10),S13,&HFFFF5BB1
FF b,c,d,a,x(k+11),S14,&H895CD7BE
FF a,b,c,d,x(k+12),S11,&H6B901122
FF d,a,b,c,x(k+13),S12,&HFD987193
FF c,d,a,b,x(k+14),S13,&HA679438E
FF b,c,d,a,x(k+15),S14,&H49B40821
GG a,b,c,d,x(k+1),S21,&HF61E2562
GG d,a,b,c,x(k+6),S22,&HC040B340
GG c,d,a,b,x(k+11),S23,&H265E5A51
GG b,c,d,a,x(k+0),S24,&HE9B6C7AA
GG a,b,c,d,x(k+5),S21,&HD62F105D
GG d,a,b,c,x(k+10),S22,&H2441453
GG c,d,a,b,x(k+15),S23,&HD8A1E681
GG b,c,d,a,x(k+4),S24,&HE7D3FBC8
GG a,b,c,d,x(k+9),S21,&H21E1CDE6
GG d,a,b,c,x(k+14),S22,&HC33707D6
GG c,d,a,b,x(k+3),S23,&HF4D50D87
GG b,c,d,a,x(k+8),S24,&H455A14ED
GG a,b,c,d,x(k+13),S21,&HA9E3E905
GG d,a,b,c,x(k+2),S22,&HFCEFA3F8
GG c,d,a,b,x(k+7),S23,&H676F02D9
GG b,c,d,a,x(k+12),S24,&H8D2A4C8A
HH a,b,c,d,x(k+5),S31,&HFFFA3942
HH d,a,b,c,x(k+8),S32,&H8771F681
HH c,d,a,b,x(k+11),S33,&H6D9D6122
HH b,c,d,a,x(k+14),S34,&HFDE5380C
HH a,b,c,d,x(k+1),S31,&HA4BEEA44
HH d,a,b,c,x(k+4),S32,&H4BDECFA9
HH c,d,a,b,x(k+7),S33,&HF6BB4B60
HH b,c,d,a,x(k+10),S34,&HBEBFBC70
HH a,b,c,d,x(k+13),S31,&H289B7EC6
HH d,a,b,c,x(k+0),S32,&HEAA127FA
HH c,d,a,b,x(k+3),S33,&HD4EF3085
HH b,c,d,a,x(k+6),S34,&H4881D05
HH a,b,c,d,x(k+9),S31,&HD9D4D039
HH d,a,b,c,x(k+12),S32,&HE6DB99E5
HH c,d,a,b,x(k+15),S33,&H1FA27CF8
HH b,c,d,a,x(k+2),S34,&HC4AC5665
II a,b,c,d,x(k+0),S41,&HF4292244
II d,a,b,c,x(k+7),S42,&H432AFF97
II c,d,a,b,x(k+14),S43,&HAB9423A7
II b,c,d,a,x(k+5),S44,&HFC93A039
II a,b,c,d,x(k+12),S41,&H655B59C3
II d,a,b,c,x(k+3),S42,&H8F0CCC92
II c,d,a,b,x(k+10),S43,&HFFEFF47D
II b,c,d,a,x(k+1),S44,&H85845DD1
II a,b,c,d,x(k+8),S41,&H6FA87E4F
II d,a,b,c,x(k+15),S42,&HFE2CE6E0
II c,d,a,b,x(k+6),S43,&HA3014314
II b,c,d,a,x(k+13),S44,&H4E0811A1
II a,b,c,d,x(k+4),S41,&HF7537E82
II d,a,b,c,x(k+11),S42,&HBD3AF235
II c,d,a,b,x(k+2),S43,&H2AD7D2BB
II b,c,d,a,x(k+9),S44,&HEB86D391
a=AddUnsigned(a,AA)
b=AddUnsigned(b,BB)
c=AddUnsigned(c,CC)
d=AddUnsigned(d,DD)
Next
MD5=LCase(WordToHex(a)&WordToHex(b)&WordToHex(c)&WordToHex(d))
End Function
End Class
Function SHA1(ByVal s)
SHA1=calcSHA1(s)
End Function
Function AndW(ByRef pBytWord1Ary,ByRef pBytWord2Ary)
Dim lBytWordAry(3)
Dim lLngIndex
For lLngIndex=0 To 3
lBytWordAry(lLngIndex)=CByte(pBytWord1Ary(lLngIndex)And pBytWord2Ary(lLngIndex))
Next
AndW=lBytWordAry
End Function
Function OrW(ByRef pBytWord1Ary,ByRef pBytWord2Ary)
Dim lBytWordAry(3)
Dim lLngIndex
For lLngIndex=0 To 3
lBytWordAry(lLngIndex)=CByte(pBytWord1Ary(lLngIndex)Or pBytWord2Ary(lLngIndex))
Next
OrW=lBytWordAry
End Function
Function XorW(ByRef pBytWord1Ary,ByRef pBytWord2Ary)
Dim lBytWordAry(3)
Dim lLngIndex
For lLngIndex=0 To 3
lBytWordAry(lLngIndex)=CByte(pBytWord1Ary(lLngIndex)Xor pBytWord2Ary(lLngIndex))
Next
XorW=lBytWordAry
End Function
Function NotW(ByRef pBytWordAry)
Dim lBytWordAry(3)
Dim lLngIndex
For lLngIndex=0 To 3
lBytWordAry(lLngIndex)=Not CByte(pBytWordAry(lLngIndex))
Next
NotW=lBytWordAry
End Function
Function AddW(ByRef pBytWord1Ary,ByRef pBytWord2Ary)
Dim lLngIndex
Dim lIntTotal
Dim lBytWordAry(3)
For lLngIndex=3 To 0 Step-1
If lLngIndex=3 Then
lIntTotal=CInt(pBytWord1Ary(lLngIndex))+pBytWord2Ary(lLngIndex)
lBytWordAry(lLngIndex)=lIntTotal Mod 256
Else
lIntTotal=CInt(pBytWord1Ary(lLngIndex))+pBytWord2Ary(lLngIndex)+(lIntTotal \ 256)
lBytWordAry(lLngIndex)=lIntTotal Mod 256
End If
Next
AddW=lBytWordAry
End Function
Function CircShiftLeftW(ByRef pBytWordAry,ByRef pLngShift)
Dim lDbl1
Dim lDbl2
lDbl1=WordToDouble(pBytWordAry)
lDbl2=lDbl1
lDbl1=CDbl(lDbl1 *(2 ^ pLngShift))
lDbl2=CDbl(lDbl2 /(2 ^(32-pLngShift)))
CircShiftLeftW=OrW(DoubleToWord(lDbl1),DoubleToWord(lDbl2))
End Function
Function WordToHex(ByRef pBytWordAry)
Dim lLngIndex
For lLngIndex=0 To 3
WordToHex=WordToHex&Right("0"& Hex(pBytWordAry(lLngIndex)),2)
Next
End Function
Function HexToWord(ByRef pStrHex)
HexToWord=DoubleToWord(CDbl("&h"&pStrHex))' needs "#" at end for VB?
End Function
Function DoubleToWord(ByRef pDblValue)
Dim lBytWordAry(3)
lBytWordAry(0)=Int(DMod(pDblValue,2 ^ 32)/(2 ^ 24))
lBytWordAry(1)=Int(DMod(pDblValue,2 ^ 24)/(2 ^ 16))
lBytWordAry(2)=Int(DMod(pDblValue,2 ^ 16)/(2 ^ 8))
lBytWordAry(3)=Int(DMod(pDblValue,2 ^ 8))
DoubleToWord=lBytWordAry
End Function
Function WordToDouble(ByRef pBytWordAry)
WordToDouble=CDbl((pBytWordAry(0)*(2 ^ 24))+(pBytWordAry(1)*(2 ^ 16))+(pBytWordAry(2)*(2 ^ 8))+pBytWordAry(3))
End Function
Function DMod(ByRef pDblValue,ByRef pDblDivisor)
Dim lDblMod
lDblMod=CDbl(CDbl(pDblValue)-(Int(CDbl(pDblValue)/ CDbl(pDblDivisor))* CDbl(pDblDivisor)))
If lDblMod<0 Then
lDblMod=CDbl(lDblMod+pDblDivisor)
End If
DMod=lDblMod
End Function
Function SHA1_F(ByRef lIntT,ByRef pBytWordBAry,ByRef pBytWordCAry,ByRef pBytWordDAry)
If lIntT<=19 Then
SHA1_F=OrW(AndW(pBytWordBAry,pBytWordCAry),AndW((NotW(pBytWordBAry)),pBytWordDAry))
ElseIf lIntT<=39 Then
SHA1_F=XorW(XorW(pBytWordBAry,pBytWordCAry),pBytWordDAry)
ElseIf lIntT<=59 Then
SHA1_F=OrW(OrW(AndW(pBytWordBAry,pBytWordCAry),AndW(pBytWordBAry,pBytWordDAry)),AndW(pBytWordCAry,pBytWordDAry))
Else
SHA1_F=XorW(XorW(pBytWordBAry,pBytWordCAry),pBytWordDAry)
End If
End Function
Function calcSHA1(ByVal pStrMessage)
Dim lLngLen
Dim lBytLenW
Dim lStrPadMessage
Dim lLngNumBlocks
Dim lVarWordWAry(79)
Dim lLngTempWordWAry
Dim lStrBlockText
Dim lStrWordText
Dim lLngBlock
Dim lIntT
Dim lBytTempAry
Dim lVarWordKAry(3)
Dim lBytWordH0Ary
Dim lBytWordH1Ary
Dim lBytWordH2Ary
Dim lBytWordH3Ary
Dim lBytWordH4Ary
Dim lBytWordAAry
Dim lBytWordBAry
Dim lBytWordCAry
Dim lBytWordDAry
Dim lBytWordEAry
Dim lBytWordFAry
lLngLen=Len(pStrMessage)
lBytLenW=DoubleToWord(CDbl(lLngLen)* 8)
lStrPadMessage=pStrMessage&Chr(128)&String((128-(lLngLen Mod 64)-9)Mod 64,Chr(0))&String(4,Chr(0))&Chr(lBytLenW(0))&Chr(lBytLenW(1))&Chr(lBytLenW(2))&Chr(lBytLenW(3))
lLngNumBlocks=Len(lStrPadMessage)/ 64
lVarWordKAry(0)=HexToWord("5A827999")
lVarWordKAry(1)=HexToWord("6ED9EBA1")
lVarWordKAry(2)=HexToWord("8F1BBCDC")
lVarWordKAry(3)=HexToWord("CA62C1D6")
lBytWordH0Ary=HexToWord("67452301")
lBytWordH1Ary=HexToWord("EFCDAB89")
lBytWordH2Ary=HexToWord("98BADCFE")
lBytWordH3Ary=HexToWord("10325476")
lBytWordH4Ary=HexToWord("C3D2E1F0")
For lLngBlock=0 To lLngNumBlocks-1
lStrBlockText=Mid(lStrPadMessage,(lLngBlock * 64)+1,64)
For lIntT=0 To 15
lStrWordText=Mid(lStrBlockText,(lIntT * 4)+1,4)
lVarWordWAry(lIntT)=Array(Asc(Mid(lStrWordText,1,1)),Asc(Mid(lStrWordText,2,1)),Asc(Mid(lStrWordText,3,1)),Asc(Mid(lStrWordText,4,1)))
Next
For lIntT=16 To 79
lVarWordWAry(lIntT)=CircShiftLeftW(XorW(XorW(XorW(lVarWordWAry(lIntT-3),lVarWordWAry(lIntT-8)),lVarWordWAry(lIntT-14)),lVarWordWAry(lIntT-16)),1)
Next
lBytWordAAry=lBytWordH0Ary
lBytWordBAry=lBytWordH1Ary
lBytWordCAry=lBytWordH2Ary
lBytWordDAry=lBytWordH3Ary
lBytWordEAry=lBytWordH4Ary
For lIntT=0 To 79
lBytWordFAry=SHA1_F(lIntT,lBytWordBAry,lBytWordCAry,lBytWordDAry)
lBytTempAry=AddW(AddW(AddW(AddW(CircShiftLeftW(lBytWordAAry,5),lBytWordFAry),lBytWordEAry),lVarWordWAry(lIntT)),lVarWordKAry(lIntT \ 20))
lBytWordEAry=lBytWordDAry
lBytWordDAry=lBytWordCAry
lBytWordCAry=CircShiftLeftW(lBytWordBAry,30)
lBytWordBAry=lBytWordAAry
lBytWordAAry=lBytTempAry
Next
lBytWordH0Ary=AddW(lBytWordH0Ary,lBytWordAAry)
lBytWordH1Ary=AddW(lBytWordH1Ary,lBytWordBAry)
lBytWordH2Ary=AddW(lBytWordH2Ary,lBytWordCAry)
lBytWordH3Ary=AddW(lBytWordH3Ary,lBytWordDAry)
lBytWordH4Ary=AddW(lBytWordH4Ary,lBytWordEAry)
Next
calcSHA1=WordToHex(lBytWordH0Ary)&WordToHex(lBytWordH1Ary)&WordToHex(lBytWordH2Ary)&WordToHex(lBytWordH3Ary)&WordToHex(lBytWordH4Ary)
calcSHA1=LCase(calcSHA1)
End Function
CLASS fwxSHA256
Private m_lOnBits(30)
Private m_l2Power(30)
Private K(63)
Private BITS_TO_A_BYTE
Private BYTES_TO_A_WORD
Private BITS_TO_A_WORD
Private Sub Class_Initialize()
BITS_TO_A_BYTE=8
BYTES_TO_A_WORD=4
BITS_TO_A_WORD=32
m_lOnBits(0)=CLng(1)
m_lOnBits(1)=CLng(3)
m_lOnBits(2)=CLng(7)
m_lOnBits(3)=CLng(15)
m_lOnBits(4)=CLng(31)
m_lOnBits(5)=CLng(63)
m_lOnBits(6)=CLng(127)
m_lOnBits(7)=CLng(255)
m_lOnBits(8)=CLng(511)
m_lOnBits(9)=CLng(1023)
m_lOnBits(10)=CLng(2047)
m_lOnBits(11)=CLng(4095)
m_lOnBits(12)=CLng(8191)
m_lOnBits(13)=CLng(16383)
m_lOnBits(14)=CLng(32767)
m_lOnBits(15)=CLng(65535)
m_lOnBits(16)=CLng(131071)
m_lOnBits(17)=CLng(262143)
m_lOnBits(18)=CLng(524287)
m_lOnBits(19)=CLng(1048575)
m_lOnBits(20)=CLng(2097151)
m_lOnBits(21)=CLng(4194303)
m_lOnBits(22)=CLng(8388607)
m_lOnBits(23)=CLng(16777215)
m_lOnBits(24)=CLng(33554431)
m_lOnBits(25)=CLng(67108863)
m_lOnBits(26)=CLng(134217727)
m_lOnBits(27)=CLng(268435455)
m_lOnBits(28)=CLng(536870911)
m_lOnBits(29)=CLng(1073741823)
m_lOnBits(30)=CLng(2147483647)
m_l2Power(0)=CLng(1)
m_l2Power(1)=CLng(2)
m_l2Power(2)=CLng(4)
m_l2Power(3)=CLng(8)
m_l2Power(4)=CLng(16)
m_l2Power(5)=CLng(32)
m_l2Power(6)=CLng(64)
m_l2Power(7)=CLng(128)
m_l2Power(8)=CLng(256)
m_l2Power(9)=CLng(512)
m_l2Power(10)=CLng(1024)
m_l2Power(11)=CLng(2048)
m_l2Power(12)=CLng(4096)
m_l2Power(13)=CLng(8192)
m_l2Power(14)=CLng(16384)
m_l2Power(15)=CLng(32768)
m_l2Power(16)=CLng(65536)
m_l2Power(17)=CLng(131072)
m_l2Power(18)=CLng(262144)
m_l2Power(19)=CLng(524288)
m_l2Power(20)=CLng(1048576)
m_l2Power(21)=CLng(2097152)
m_l2Power(22)=CLng(4194304)
m_l2Power(23)=CLng(8388608)
m_l2Power(24)=CLng(16777216)
m_l2Power(25)=CLng(33554432)
m_l2Power(26)=CLng(67108864)
m_l2Power(27)=CLng(134217728)
m_l2Power(28)=CLng(268435456)
m_l2Power(29)=CLng(536870912)
m_l2Power(30)=CLng(1073741824)
K(0)=&H428A2F98
K(1)=&H71374491
K(2)=&HB5C0FBCF
K(3)=&HE9B5DBA5
K(4)=&H3956C25B
K(5)=&H59F111F1
K(6)=&H923F82A4
K(7)=&HAB1C5ED5
K(8)=&HD807AA98
K(9)=&H12835B01
K(10)=&H243185BE
K(11)=&H550C7DC3
K(12)=&H72BE5D74
K(13)=&H80DEB1FE
K(14)=&H9BDC06A7
K(15)=&HC19BF174
K(16)=&HE49B69C1
K(17)=&HEFBE4786
K(18)=&HFC19DC6
K(19)=&H240CA1CC
K(20)=&H2DE92C6F
K(21)=&H4A7484AA
K(22)=&H5CB0A9DC
K(23)=&H76F988DA
K(24)=&H983E5152
K(25)=&HA831C66D
K(26)=&HB00327C8
K(27)=&HBF597FC7
K(28)=&HC6E00BF3
K(29)=&HD5A79147
K(30)=&H6CA6351
K(31)=&H14292967
K(32)=&H27B70A85
K(33)=&H2E1B2138
K(34)=&H4D2C6DFC
K(35)=&H53380D13
K(36)=&H650A7354
K(37)=&H766A0ABB
K(38)=&H81C2C92E
K(39)=&H92722C85
K(40)=&HA2BFE8A1
K(41)=&HA81A664B
K(42)=&HC24B8B70
K(43)=&HC76C51A3
K(44)=&HD192E819
K(45)=&HD6990624
K(46)=&HF40E3585
K(47)=&H106AA070
K(48)=&H19A4C116
K(49)=&H1E376C08
K(50)=&H2748774C
K(51)=&H34B0BCB5
K(52)=&H391C0CB3
K(53)=&H4ED8AA4A
K(54)=&H5B9CCA4F
K(55)=&H682E6FF3
K(56)=&H748F82EE
K(57)=&H78A5636F
K(58)=&H84C87814
K(59)=&H8CC70208
K(60)=&H90BEFFFA
K(61)=&HA4506CEB
K(62)=&HBEF9A3F7
K(63)=&HC67178F2
End Sub
Private Function LShift(lValue,iShiftBits)
If iShiftBits=0 Then
LShift=lValue
Exit Function
ElseIf iShiftBits=31 Then
If lValue And 1 Then
LShift=&H80000000
Else
LShift=0
End If
Exit Function
ElseIf iShiftBits<0 Or iShiftBits>31 Then
Err.Raise 6
End If
If(lValue And m_l2Power(31-iShiftBits))Then
LShift=((lValue And m_lOnBits(31-(iShiftBits+1)))* m_l2Power(iShiftBits))Or &H80000000
Else
LShift=((lValue And m_lOnBits(31-iShiftBits))* m_l2Power(iShiftBits))
End If
End Function
Private Function RShift(lValue,iShiftBits)
If iShiftBits=0 Then
RShift=lValue
Exit Function
ElseIf iShiftBits=31 Then
If lValue And &H80000000 Then
RShift=1
Else
RShift=0
End If
Exit Function
ElseIf iShiftBits<0 Or iShiftBits>31 Then
Err.Raise 6
End If
RShift=(lValue And &H7FFFFFFE)\ m_l2Power(iShiftBits)
If(lValue And &H80000000)Then
RShift=(RShift Or(&H40000000 \ m_l2Power(iShiftBits-1)))
End If
End Function
Private Function AddUnsigned(lX,lY)
Dim lX4
Dim lY4
Dim lX8
Dim lY8
Dim lResult
lX8=lX And &H80000000
lY8=lY And &H80000000
lX4=lX And &H40000000
lY4=lY And &H40000000
lResult=(lX And &H3FFFFFFF)+(lY And &H3FFFFFFF)
If lX4 And lY4 Then
lResult=lResult Xor &H80000000 Xor lX8 Xor lY8
ElseIf lX4 Or lY4 Then
If lResult And &H40000000 Then
lResult=lResult Xor &HC0000000 Xor lX8 Xor lY8
Else
lResult=lResult Xor &H40000000 Xor lX8 Xor lY8
End If
Else
lResult=lResult Xor lX8 Xor lY8
End If
AddUnsigned=lResult
End Function
Private Function Ch(x,y,z)
Ch=((x And y)Xor((Not x)And z))
End Function
Private Function Maj(x,y,z)
Maj=((x And y)Xor(x And z)Xor(y And z))
End Function
Private Function S(x,n)
S=(RShift(x,(n And m_lOnBits(4)))Or LShift(x,(32-(n And m_lOnBits(4)))))
End Function
Private Function R(x,n)
R=RShift(x,CInt(n And m_lOnBits(4)))
End Function
Private Function Sigma0(x)
Sigma0=(S(x,2)Xor S(x,13)Xor S(x,22))
End Function
Private Function Sigma1(x)
Sigma1=(S(x,6)Xor S(x,11)Xor S(x,25))
End Function
Private Function Gamma0(x)
Gamma0=(S(x,7)Xor S(x,18)Xor R(x,3))
End Function
Private Function Gamma1(x)
Gamma1=(S(x,17)Xor S(x,19)Xor R(x,10))
End Function
Private Function ConvertToWordArray(sMessage)
Dim lMessageLength
Dim lNumberOfWords
Dim lWordArray()
Dim lBytePosition
Dim lByteCount
Dim lWordCount
Dim lByte
Const MODULUS_BITS=512
Const CONGRUENT_BITS=448
lMessageLength=Len(sMessage)
lNumberOfWords=(((lMessageLength+((MODULUS_BITS-CONGRUENT_BITS)\ BITS_TO_A_BYTE))\(MODULUS_BITS \ BITS_TO_A_BYTE))+1)*(MODULUS_BITS \ BITS_TO_A_WORD)
ReDim lWordArray(lNumberOfWords-1)
lBytePosition=0
lByteCount=0
Do Until lByteCount>=lMessageLength
lWordCount=lByteCount \ BYTES_TO_A_WORD
lBytePosition=(3-(lByteCount Mod BYTES_TO_A_WORD))* BITS_TO_A_BYTE
lByte=AscB(Mid(sMessage,lByteCount+1,1))
lWordArray(lWordCount)=lWordArray(lWordCount)Or LShift(lByte,lBytePosition)
lByteCount=lByteCount+1
Loop
lWordCount=lByteCount \ BYTES_TO_A_WORD
lBytePosition=(3-(lByteCount Mod BYTES_TO_A_WORD))* BITS_TO_A_BYTE
lWordArray(lWordCount)=lWordArray(lWordCount)Or LShift(&H80,lBytePosition)
lWordArray(lNumberOfWords-1)=LShift(lMessageLength,3)
lWordArray(lNumberOfWords-2)=RShift(lMessageLength,29)
ConvertToWordArray=lWordArray
End Function
Public Function SHA256(sMessage)
Dim HASH(7)
Dim M
Dim W(63)
Dim a
Dim b
Dim c
Dim d
Dim e
Dim f
Dim g
Dim h
Dim i
Dim j
Dim T1
Dim T2
HASH(0)=&H6A09E667
HASH(1)=&HBB67AE85
HASH(2)=&H3C6EF372
HASH(3)=&HA54FF53A
HASH(4)=&H510E527F
HASH(5)=&H9B05688C
HASH(6)=&H1F83D9AB
HASH(7)=&H5BE0CD19
M=ConvertToWordArray(sMessage)
For i=0 To UBound(M)Step 16
a=HASH(0)
b=HASH(1)
c=HASH(2)
d=HASH(3)
e=HASH(4)
f=HASH(5)
g=HASH(6)
h=HASH(7)
For j=0 To 63
If j<16 Then
W(j)=M(j+i)
Else
W(j)=AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j-2)),W(j-7)),Gamma0(W(j-15))),W(j-16))
End If
T1=AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h,Sigma1(e)),Ch(e,f,g)),K(j)),W(j))
T2=AddUnsigned(Sigma0(a),Maj(a,b,c))
h=g
g=f
f=e
e=AddUnsigned(d,T1)
d=c
c=b
b=a
a=AddUnsigned(T1,T2)
Next
HASH(0)=AddUnsigned(a,HASH(0))
HASH(1)=AddUnsigned(b,HASH(1))
HASH(2)=AddUnsigned(c,HASH(2))
HASH(3)=AddUnsigned(d,HASH(3))
HASH(4)=AddUnsigned(e,HASH(4))
HASH(5)=AddUnsigned(f,HASH(5))
HASH(6)=AddUnsigned(g,HASH(6))
HASH(7)=AddUnsigned(h,HASH(7))
Next
SHA256=LCase(Right("00000000"& Hex(HASH(0)),8)&Right("00000000"& Hex(HASH(1)),8)&Right("00000000"& Hex(HASH(2)),8)&Right("00000000"& Hex(HASH(3)),8)&Right("00000000"& Hex(HASH(4)),8)&Right("00000000"& Hex(HASH(5)),8)&Right("00000000"& Hex(HASH(6)),8)&Right("00000000"& Hex(HASH(7)),8))
End Function
END CLASS
Class cls_SagePay300_ClassicASP
Public Password
Public Function DecodeAndDecrypt(ByVal strIn)
DecodeAndDecrypt=AESDecrypt(mid(strIn,2),Me.Password)
End Function
Public Function EncryptAndEncode(ByVal strIn)
EncryptAndEncode="@"&AESEncrypt(strIn,Me.Password)
End Function
Private m_lOnBits(30)
Private m_l2Power(30)
Private m_bytOnBits(7)
Private m_byt2Power(7)
Private m_InCo(3)
Private m_fbsub(255)
Private m_rbsub(255)
Private m_ptab(255)
Private m_ltab(255)
Private m_ftable(255)
Private m_rtable(255)
Private m_rco(29)
Private m_Nk
Private m_Nb
Private m_Nr
Private m_fi(23)
Private m_ri(23)
Private m_fkey(119)
Private m_rkey(119)
Private Sub Class_Initialize()
m_InCo(0)=&HB
m_InCo(1)=&HD
m_InCo(2)=&H9
m_InCo(3)=&HE
m_bytOnBits(0)=1
m_bytOnBits(1)=3
m_bytOnBits(2)=7
m_bytOnBits(3)=15
m_bytOnBits(4)=31
m_bytOnBits(5)=63
m_bytOnBits(6)=127
m_bytOnBits(7)=255
m_byt2Power(0)=1
m_byt2Power(1)=2
m_byt2Power(2)=4
m_byt2Power(3)=8
m_byt2Power(4)=16
m_byt2Power(5)=32
m_byt2Power(6)=64
m_byt2Power(7)=128
m_lOnBits(0)=1
m_lOnBits(1)=3
m_lOnBits(2)=7
m_lOnBits(3)=15
m_lOnBits(4)=31
m_lOnBits(5)=63
m_lOnBits(6)=127
m_lOnBits(7)=255
m_lOnBits(8)=511
m_lOnBits(9)=1023
m_lOnBits(10)=2047
m_lOnBits(11)=4095
m_lOnBits(12)=8191
m_lOnBits(13)=16383
m_lOnBits(14)=32767
m_lOnBits(15)=65535
m_lOnBits(16)=131071
m_lOnBits(17)=262143
m_lOnBits(18)=524287
m_lOnBits(19)=1048575
m_lOnBits(20)=2097151
m_lOnBits(21)=4194303
m_lOnBits(22)=8388607
m_lOnBits(23)=16777215
m_lOnBits(24)=33554431
m_lOnBits(25)=67108863
m_lOnBits(26)=134217727
m_lOnBits(27)=268435455
m_lOnBits(28)=536870911
m_lOnBits(29)=1073741823
m_lOnBits(30)=2147483647
m_l2Power(0)=1
m_l2Power(1)=2
m_l2Power(2)=4
m_l2Power(3)=8
m_l2Power(4)=16
m_l2Power(5)=32
m_l2Power(6)=64
m_l2Power(7)=128
m_l2Power(8)=256
m_l2Power(9)=512
m_l2Power(10)=1024
m_l2Power(11)=2048
m_l2Power(12)=4096
m_l2Power(13)=8192
m_l2Power(14)=16384
m_l2Power(15)=32768
m_l2Power(16)=65536
m_l2Power(17)=131072
m_l2Power(18)=262144
m_l2Power(19)=524288
m_l2Power(20)=1048576
m_l2Power(21)=2097152
m_l2Power(22)=4194304
m_l2Power(23)=8388608
m_l2Power(24)=16777216
m_l2Power(25)=33554432
m_l2Power(26)=67108864
m_l2Power(27)=134217728
m_l2Power(28)=268435456
m_l2Power(29)=536870912
m_l2Power(30)=1073741824
End Sub
Private Function LShift(lValue,iShiftBits)
If iShiftBits=0 Then
LShift=lValue
Exit Function
ElseIf iShiftBits=31 Then
If lValue And 1 Then
LShift=&H80000000
Else
LShift=0
End If
Exit Function
ElseIf iShiftBits<0 Or iShiftBits>31 Then
Err.Raise 6
End If
If(lValue And m_l2Power(31-iShiftBits))Then
LShift=((lValue And m_lOnBits(31-(iShiftBits+1)))* m_l2Power(iShiftBits))Or &H80000000
Else
LShift=((lValue And m_lOnBits(31-iShiftBits))* m_l2Power(iShiftBits))
End If
End Function
Private Function RShift(lValue,iShiftBits)
If iShiftBits=0 Then
RShift=lValue
Exit Function
ElseIf iShiftBits=31 Then
If lValue And &H80000000 Then
RShift=1
Else
RShift=0
End If
Exit Function
ElseIf iShiftBits<0 Or iShiftBits>31 Then
Err.Raise 6
End If
RShift=(lValue And &H7FFFFFFE)\ m_l2Power(iShiftBits)
If(lValue And &H80000000)Then
RShift=(RShift Or(&H40000000 \ m_l2Power(iShiftBits-1)))
End If
End Function
Private Function LShiftByte(bytValue,bytShiftBits)
If bytShiftBits=0 Then
LShiftByte=bytValue
Exit Function
ElseIf bytShiftBits=7 Then
If bytValue And 1 Then
LShiftByte=&H80
Else
LShiftByte=0
End If
Exit Function
ElseIf bytShiftBits<0 Or bytShiftBits>7 Then
Err.Raise 6
End If
LShiftByte=((bytValue And m_bytOnBits(7-bytShiftBits))* m_byt2Power(bytShiftBits))
End Function
Private Function RShiftByte(bytValue,bytShiftBits)
If bytShiftBits=0 Then
RShiftByte=bytValue
Exit Function
ElseIf bytShiftBits=7 Then
If bytValue And &H80 Then
RShiftByte=1
Else
RShiftByte=0
End If
Exit Function
ElseIf bytShiftBits<0 Or bytShiftBits>7 Then
Err.Raise 6
End If
RShiftByte=bytValue \ m_byt2Power(bytShiftBits)
End Function
Private Function RotateLeft(lValue,iShiftBits)
RotateLeft=LShift(lValue,iShiftBits)Or RShift(lValue,(32-iShiftBits))
End Function
Private Function RotateLeftByte(bytValue,bytShiftBits)
RotateLeftByte=LShiftByte(bytValue,bytShiftBits)Or RShiftByte(bytValue,(8-bytShiftBits))
End Function
Private Function Pack(b())
Dim lCount
Dim lTemp
For lCount=0 To 3
lTemp=b(lCount)
Pack=Pack Or LShift(lTemp,(lCount * 8))
Next
End Function
Private Function PackFrom(b(),k)
Dim lCount
Dim lTemp
For lCount=0 To 3
lTemp=b(lCount+k)
PackFrom=PackFrom Or LShift(lTemp,(lCount * 8))
Next
End Function
Private Sub Unpack(a,b())
b(0)=a And m_lOnBits(7)
b(1)=RShift(a,8)And m_lOnBits(7)
b(2)=RShift(a,16)And m_lOnBits(7)
b(3)=RShift(a,24)And m_lOnBits(7)
End Sub
Private Sub UnpackFrom(a,b(),k)
b(0+k)=a And m_lOnBits(7)
b(1+k)=RShift(a,8)And m_lOnBits(7)
b(2+k)=RShift(a,16)And m_lOnBits(7)
b(3+k)=RShift(a,24)And m_lOnBits(7)
End Sub
Private Function xtime(a)
Dim b
If(a And &H80)Then
b=&H1B
Else
b=0
End If
xtime=LShiftByte(a,1)
xtime=xtime Xor b
End Function
Private Function bmul(x,y)
If x<>0 And y<>0 Then
bmul=m_ptab((CLng(m_ltab(x))+CLng(m_ltab(y)))Mod 255)
Else
bmul=0
End If
End Function
Private Function SubByte(a)
Dim b(3)
Unpack a,b
b(0)=m_fbsub(b(0))
b(1)=m_fbsub(b(1))
b(2)=m_fbsub(b(2))
b(3)=m_fbsub(b(3))
SubByte=Pack(b)
End Function
Private Function product(x,y)
Dim xb(3)
Dim yb(3)
Unpack x,xb
Unpack y,yb
product=bmul(xb(0),yb(0))Xor bmul(xb(1),yb(1))Xor bmul(xb(2),yb(2))Xor bmul(xb(3),yb(3))
End Function
Private Function InvMixCol(x)
Dim y
Dim m
Dim b(3)
m=Pack(m_InCo)
b(3)=product(m,x)
m=RotateLeft(m,24)
b(2)=product(m,x)
m=RotateLeft(m,24)
b(1)=product(m,x)
m=RotateLeft(m,24)
b(0)=product(m,x)
y=Pack(b)
InvMixCol=y
End Function
Private Function ByteSub(x)
Dim y
Dim z
z=x
y=m_ptab(255-m_ltab(z))
z=y
z=RotateLeftByte(z,1)
y=y Xor z
z=RotateLeftByte(z,1)
y=y Xor z
z=RotateLeftByte(z,1)
y=y Xor z
z=RotateLeftByte(z,1)
y=y Xor z
y=y Xor &H63
ByteSub=y
End Function
Public Sub gentables()
Dim i
Dim y
Dim b(3)
Dim ib
m_ltab(0)=0
m_ptab(0)=1
m_ltab(1)=0
m_ptab(1)=3
m_ltab(3)=1
For i=2 To 255
m_ptab(i)=m_ptab(i-1)Xor xtime(m_ptab(i-1))
m_ltab(m_ptab(i))=i
Next
m_fbsub(0)=&H63
m_rbsub(&H63)=0
For i=1 To 255
ib=i
y=ByteSub(ib)
m_fbsub(i)=y
m_rbsub(y)=i
Next
y=1
For i=0 To 29
m_rco(i)=y
y=xtime(y)
Next
For i=0 To 255
y=m_fbsub(i)
b(3)=y Xor xtime(y)
b(2)=y
b(1)=y
b(0)=xtime(y)
m_ftable(i)=Pack(b)
y=m_rbsub(i)
b(3)=bmul(m_InCo(0),y)
b(2)=bmul(m_InCo(1),y)
b(1)=bmul(m_InCo(2),y)
b(0)=bmul(m_InCo(3),y)
m_rtable(i)=Pack(b)
Next
End Sub
Public Sub gkey(nb,nk,key())
Dim i
Dim j
Dim k
Dim m
Dim N
Dim C1
Dim C2
Dim C3
Dim CipherKey(7)
m_Nb=nb
m_Nk=nk
If m_Nb>=m_Nk Then
m_Nr=6+m_Nb
Else
m_Nr=6+m_Nk
End If
C1=1
If m_Nb<8 Then
C2=2
C3=3
Else
C2=3
C3=4
End If
For j=0 To nb-1
m=j * 3
m_fi(m)=(j+C1)Mod nb
m_fi(m+1)=(j+C2)Mod nb
m_fi(m+2)=(j+C3)Mod nb
m_ri(m)=(nb+j-C1)Mod nb
m_ri(m+1)=(nb+j-C2)Mod nb
m_ri(m+2)=(nb+j-C3)Mod nb
Next
N=m_Nb *(m_Nr+1)
For i=0 To m_Nk-1
j=i * 4
CipherKey(i)=PackFrom(key,j)
Next
For i=0 To m_Nk-1
m_fkey(i)=CipherKey(i)
Next
j=m_Nk
k=0
Do While j<N
m_fkey(j)=m_fkey(j-m_Nk)Xor SubByte(RotateLeft(m_fkey(j-1),24))Xor m_rco(k)
If m_Nk<=6 Then
i=1
Do While i<m_Nk And(i+j)<N
m_fkey(i+j)=m_fkey(i+j-m_Nk)Xor m_fkey(i+j-1)
i=i+1
Loop
Else
i=1
Do While i<4 And(i+j)<N
m_fkey(i+j)=m_fkey(i+j-m_Nk)Xor m_fkey(i+j-1)
i=i+1
Loop
If j+4<N Then
m_fkey(j+4)=m_fkey(j+4-m_Nk)Xor SubByte(m_fkey(j+3))
End If
i=5
Do While i<m_Nk And(i+j)<N
m_fkey(i+j)=m_fkey(i+j-m_Nk)Xor m_fkey(i+j-1)
i=i+1
Loop
End If
j=j+m_Nk
k=k+1
Loop
For j=0 To m_Nb-1
m_rkey(j+N-nb)=m_fkey(j)
Next
i=m_Nb
Do While i<N-m_Nb
k=N-m_Nb-i
For j=0 To m_Nb-1
m_rkey(k+j)=InvMixCol(m_fkey(i+j))
Next
i=i+m_Nb
Loop
j=N-m_Nb
Do While j<N
m_rkey(j-N+m_Nb)=m_fkey(j)
j=j+1
Loop
End Sub
Public Sub encrypt(buff())
Dim i
Dim j
Dim k
Dim m
Dim a(7)
Dim b(7)
Dim x
Dim y
Dim t
For i=0 To m_Nb-1
j=i * 4
a(i)=PackFrom(buff,j)
a(i)=a(i)Xor m_fkey(i)
Next
k=m_Nb
x=a
y=b
For i=1 To m_Nr-1
For j=0 To m_Nb-1
m=j * 3
y(j)=m_fkey(k)Xor m_ftable(x(j)And m_lOnBits(7))Xor RotateLeft(m_ftable(RShift(x(m_fi(m)),8)And m_lOnBits(7)),8)Xor RotateLeft(m_ftable(RShift(x(m_fi(m+1)),16)And m_lOnBits(7)),16)Xor RotateLeft(m_ftable(RShift(x(m_fi(m+2)),24)And m_lOnBits(7)),24)
k=k+1
Next
t=x
x=y
y=t
Next
For j=0 To m_Nb-1
m=j * 3
y(j)=m_fkey(k)Xor m_fbsub(x(j)And m_lOnBits(7))Xor RotateLeft(m_fbsub(RShift(x(m_fi(m)),8)And m_lOnBits(7)),8)Xor RotateLeft(m_fbsub(RShift(x(m_fi(m+1)),16)And m_lOnBits(7)),16)Xor RotateLeft(m_fbsub(RShift(x(m_fi(m+2)),24)And m_lOnBits(7)),24)
k=k+1
Next
For i=0 To m_Nb-1
j=i * 4
UnpackFrom y(i),buff,j
x(i)=0
y(i)=0
Next
End Sub
Public Sub decrypt(buff())
Dim i
Dim j
Dim k
Dim m
Dim a(7)
Dim b(7)
Dim x
Dim y
Dim t
For i=0 To m_Nb-1
j=i * 4
a(i)=PackFrom(buff,j)
a(i)=a(i)Xor m_rkey(i)
Next
k=m_Nb
x=a
y=b
For i=1 To m_Nr-1
For j=0 To m_Nb-1
m=j * 3
y(j)=m_rkey(k)Xor m_rtable(x(j)And m_lOnBits(7))Xor RotateLeft(m_rtable(RShift(x(m_ri(m)),8)And m_lOnBits(7)),8)Xor RotateLeft(m_rtable(RShift(x(m_ri(m+1)),16)And m_lOnBits(7)),16)Xor RotateLeft(m_rtable(RShift(x(m_ri(m+2)),24)And m_lOnBits(7)),24)
k=k+1
Next
t=x
x=y
y=t
Next
For j=0 To m_Nb-1
m=j * 3
y(j)=m_rkey(k)Xor m_rbsub(x(j)And m_lOnBits(7))Xor RotateLeft(m_rbsub(RShift(x(m_ri(m)),8)And m_lOnBits(7)),8)Xor RotateLeft(m_rbsub(RShift(x(m_ri(m+1)),16)And m_lOnBits(7)),16)Xor RotateLeft(m_rbsub(RShift(x(m_ri(m+2)),24)And m_lOnBits(7)),24)
k=k+1
Next
For i=0 To m_Nb-1
j=i * 4
UnpackFrom y(i),buff,j
x(i)=0
y(i)=0
Next
End Sub
Private Function IsInitialized(vArray)
On Error Resume Next
IsInitialized=IsNumeric(UBound(vArray))
End Function
Private Sub CopyBytesASP(bytDest,lDestStart,bytSource(),lSourceStart,lLength)
Dim lCount
lCount=0
Do
bytDest(lDestStart+lCount)=bytSource(lSourceStart+lCount)
lCount=lCount+1
Loop Until lCount=lLength
End Sub
public Sub XORBlock(bytData1,bytData2)
Dim lCount
for lCount=0 to 15
bytData1(lCount)=(bytData1(lCount)XOR bytData2(lCount))
next
end Sub
Public Function EncryptData(bytMessage,bytPassword)
Dim bytKey(15)
Dim bytIn()
Dim bytOut()
Dim bytLast(15)
Dim bytTemp(15)
Dim lCount
Dim lLength
Dim lEncodedLength
Dim bytLen(3)
Dim lPosition
If Not IsInitialized(bytMessage)Then
Exit Function
End If
If Not IsInitialized(bytPassword)Then
Exit Function
End If
For lCount=0 To UBound(bytPassword)
bytKey(lCount)=bytPassword(lCount)
If lCount=15 Then
Exit For
End If
Next
gentables
gkey 4,4,bytKey
lLength=UBound(bytMessage)+1
lEncodedLength=lLength
If lEncodedLength Mod 16<>0 Then
lEncodedLength=lEncodedLength+16-(lEncodedLength Mod 16)
End If
ReDim bytIn(lEncodedLength-1)
ReDim bytOut(lEncodedLength-1)
CopyBytesASP bytLast,0,bytPassword,0,16
Unpack lLength,bytIn
CopyBytesASP bytIn,0,bytMessage,0,lLength
For lCount=0 To lEncodedLength-1 Step 16
CopyBytesASP bytTemp,0,bytIn,lCount,16
XORBlock bytTemp,bytLast
Encrypt bytTemp
CopyBytesASP bytOut,lCount,bytTemp,0,16
CopyBytesASP bytLast,0,bytTemp,0,16
Next
EncryptData=bytOut
End Function
Public Function DecryptData(bytIn,bytPassword)
Dim bytMessage()
Dim bytKey(15)
Dim bytOut()
Dim bytTemp(15)
Dim bytLast(15)
Dim lCount
Dim lLength
Dim lEncodedLength
Dim bytLen(3)
Dim lPosition
If Not IsInitialized(bytIn)Then
Exit Function
End If
If Not IsInitialized(bytPassword)Then
Exit Function
End If
lEncodedLength=UBound(bytIn)+1
If lEncodedLength Mod 16<>0 Then
Exit Function
End If
For lCount=0 To UBound(bytPassword)
bytKey(lCount)=bytPassword(lCount)
If lCount=15 Then
Exit For
End If
Next
gentables
gkey 4,4,bytKey
CopyBytesASP bytLast,0,bytPassword,0,16
ReDim bytOut(lEncodedLength-1)
For lCount=0 To lEncodedLength-1 Step 16
CopyBytesASP bytTemp,0,bytIn,lCount,16
Decrypt bytTemp
XORBlock bytTemp,bytLast
CopyBytesASP bytLast,0,bytIn,lCount,16
CopyBytesASP bytOut,lCount,bytTemp,0,16
Next
lLength=ubound(bytOut)
ReDim bytMessage(lLength)
CopyBytesASP bytMessage,0,bytOut,0,lLength+1
DecryptData=bytMessage
End Function
Function AESEncrypt(sPlain,sPassword)
Dim bytIn()
Dim bytOut
Dim bytPassword()
Dim lCount
Dim lLength
Dim sTemp
Dim lPadLength
lLength=Len(sPlain)
lPadLength=16-(lLength mod 16)
for lCount=1 to lPadLength
sPlain=sPlain&Chr(lPadLength)
next
lLength=Len(sPlain)
ReDim bytIn(lLength-1)
For lCount=1 To lLength
bytIn(lCount-1)=CByte(AscB(Mid(sPlain,lCount,1)))
Next
lLength=Len(sPassword)
ReDim bytPassword(lLength-1)
For lCount=1 To lLength
bytPassword(lCount-1)=CByte(AscB(Mid(sPassword,lCount,1)))
Next
bytOut=EncryptData(bytIn,bytPassword)
sTemp=""
For lCount=0 To UBound(bytOut)
sTemp=sTemp&Right("0"& Hex(bytOut(lCount)),2)
Next
AESEncrypt=sTemp
End Function
Function AESDecrypt(sCypher,sPassword)
Dim bytIn()
Dim bytOut
Dim bytPassword()
Dim lCount
Dim lLength
Dim sTemp
lLength=Len(sCypher)
ReDim bytIn(lLength/2-1)
For lCount=0 To lLength/2-1
bytIn(lCount)=CByte("&H"&Mid(sCypher,lCount*2+1,2))
Next
lLength=Len(sPassword)
ReDim bytPassword(lLength-1)
For lCount=1 To lLength
bytPassword(lCount-1)=CByte(AscB(Mid(sPassword,lCount,1)))
Next
bytOut=DecryptData(bytIn,bytPassword)
lLength=UBound(bytOut)+1-bytOut(UBound(bytOut))
sTemp=""
For lCount=0 To lLength-1
sTemp=sTemp&Chr(bytOut(lCount))
Next
AESDecrypt=sTemp
End Function
End Class
Function SagePay300_Crypto()
If Not IsObject(j("SagePayCryto"))Then
Set j("SagePayCryto")=New cls_SagePay300_ClassicASP
End If
Set SagePay300_Crypto=j("SagePayCryto")
End Function
Function SagePay300_AESEncrypt(ByVal strIn,ByVal strEncryptionPassword)
With SagePay300_Crypto
.Password=strEncryptionPassword
SagePay300_AESEncrypt=.EncryptAndEncode(strIn)
End With
End Function
Function SagePay300_AESDecrypt(ByVal strIn,ByVal strEncryptionPassword)
With SagePay300_Crypto
.Password=strEncryptionPassword
SagePay300_AESDecrypt=.DecodeAndDecrypt(strIn)
End With
End Function
Class QuickString
Dim stringArray,growthRate,numItems
Public Property Get Count():Count=numItems:End Property
Public Property Get Length():Length=numItems:End Property
Private MARKER
Dim Before,Divider,After
Private Sub Class_Terminate()
Erase stringArray
End Sub
Private Sub Class_Initialize()
MARKER="!%="&Timer()&":%,"
growthRate=50
numItems=0
ReDim stringArray(growthRate)
End Sub
Public Function Capacity()
On Error Resume Next
Capacity=UBound(stringArray)
If Err.Number<>0 Then
Err.Clear()
Capacity=0
End If
End Function
Public Sub Add(ByVal strValue)
strValue=strValue&""
If strValue<>"" Then
Dim curtSize
curtSize=Capacity
If curtSize=0 Then
ReDim Preserve stringArray(numItemse)
curtSize=Capacity
End If
If numItems>=curtSize Then
ReDim Preserve stringArray(Capacity+growthRate)
End If
stringArray(numItems)=strValue:numItems=numItems+1
End If
End Sub
Public Sub Reset
Erase stringArray
Call Class_Initialize()
End Sub
Public Sub Clear
Call Reset
End Sub
Public Property Let String(ByVal strValue)
Call Reset
Me.Add strValue
End Property
Public Default Property Get String()
String=Me.JoinString(Me.Divider)
End Property
Public Property Get GetString()
GetString=Me.String()
End Property
Public Property Get JoinString(ByVal d)
Dim o:o=""
Redim Preserve stringArray(numItems)
If UBound(stringArray)>0 Then
o=Join(stringArray,MARKER)
o=Left(o,Len(o)-Len(MARKER))
o=Replace(o,MARKER,d)
o=Me.Before&o&Me.After
End If
JoinString=o
End Property
Public Property Get IsSet():IsSet=CBool(numItems>0):End Property
Public Property Get IsEmpty():IsEmpty=Not IsSet:End Property
Public Property Get Has(ByVal sVal)
Dim sFullString
sFullString=MARKER&Join(stringArray,MARKER)&MARKER
Has=CBool(InStr(1,LCase(sFullString),LCase(MARKER&sVal&MARKER))>0)
sFullString=""
End Property
End Class
Class fwxDictionary
Public d
Private Sub Class_Initialize()
Set d=Dictionary()
End Sub
Private Sub Class_Terminate()
Set d=Nothing
End Sub
Public Property Let Properties(ByVal s)
UpdateCollection Me.d,s
End Property
Public Property Let Config(ByVal s)
Me.Properties=s
End Property
Public Default Property Get Item(ByVal s)
Item=d(s)
End Property
Public Property Let Item(ByVal s,ByVal v)
d(s)=v
End Property
Public Property Set Item(ByVal s,ByVal o)
Set d(s)=o
End Property
End Class
Dim DBG
Set DBG=New fwxDebugger
Class fwxDebugger
Function Setup()
Dim d
d=rxSelect(rxSelect(Request.QueryString,"debug\=\w+"),"\w+$")
If d<>"" Then
If OneWayEncrypt(d)="191f42bcb3f8095f8969c5cf5824a01488a36fba939d414320f0cc2b7307a64a" Then
DIE PrintCollection(Application.Contents)&"<br/<hr/><br/>"&Me.Snapshot
End If
If d="diego" Then
Me.Active=True
Me.Show=True
RHTM
End If
End If
End Function
Public Warnings
Public Active
Public Show
Public File
Public Count
Public History
Public Start
Public ThisStep
Public LastStep
Public Ellapsed
Public WaitingToPrint
Public Filter
Private Sub Class_Initialize()
Me.Start=Timer()
Me.ThisStep=Me.Start
Me.Active=False
Set Me.History=New QuickString
Set Me.Warnings=New Quickstring:Me.Warnings.Divider=vbCRLf
Call Me.Setup()
Debug "/*--- START DEBUGGER ---*/"
End Sub
Private LastCheckpoint
Private ThisCheckpoint
Function Checkpoint()
LastCheckpoint=IIf(ThisCheckpoint<>"",ThisCheckpoint,Me.Start)
ThisCheckpoint=Timer()
Checkpoint=Left(FormatNumber(ThisCheckpoint-LastCheckpoint,3),9)
End Function
Private Sub Class_Terminate()
If Me.Show Then
Me.Debug "TERMINATING REQUEST"
Me.Debug "..................."
End If
If IsLocal Or(IsTest Or IsAdmin)Then
If IsBack And IsDesktop And IsLoggedIn Then
If Query.EXT="" Or rxTest(Query.EXT,"asp|js")Then
If IsTrue(Me.Silent)Or IsTrue(IsMobileDevice)Or IsTrue(ColItem(j,"no-debug"))Then
ElseIf FWX_ResponseType<>"html" And(IsTest Or IsAjax)Then
ElseIf Query.FQ("autocomplete,nolog,silent")<>"" Or Query.String="" Then
Else
Response.Flush()
If Query.EXT="" Or rxTest(Query.EXT,"asp|html?")Then:Else:RTXT:End If
Dim le:le=LastErrorInfo
RW vbCRLf&vbCRLf&vbCRLf&vbCRLf&vbCRLf&vbCRLf
RW le&vbCRLf
RW "<div class='debug-query' style='"&IIf(IsTest or le<>"","","display:none;")&"'>"&Me.MiniSnapshot&"</div>"
Response.Flush()
End If
End If
End If
End If
If IsObject(j)Then
Dim f,i,o,l,os
If Me.Show Then
l=0
i=0
o=0
os=""
For Each f In j
i=i+1
If IsObject(j(f))Then
o=o+1
os=os&f&";"
Else
l=l+Len(j(f))
End If
Next
w "/*--- KILL FWX REQUEST -----"
w "FWX-REQUEST-ITEMS="&i
w "FWX-REQUEST-DATA LENGTH="&l
w "FWX-REQUEST-OBJECTS KILLED="&o&" ("&os&")"
For Each f In Split("FolderExists,FileExists,GetFile,ReadFile",",")
w "FWX-REQUEST-FSO."&f&"="&VOZ(j("_STAT-FSO."&f))
Next
w "----- KILL FWX REQUEST ---*/"
End If
For Each f In j
If IsObject(j(f))Then
os=os&f&";"
Set j(f)=Nothing
Else
j(f)=""
End If
Next
Set j=Nothing
End If
Dim FWX_QueryCost
FWX_QueryCost=Timer()-Me.Start
Application.Contents("_STAT-PI-count")=VOZ(Application.Contents("_STAT-PI-count"))+1
Application.Contents("_STAT-PI-delay")=FormatNumber(VOZ(Application.Contents("_STAT-PI-delay"))+FWX_QueryCost,2)
Application.Contents("_STAT-PI-score")=FormatNumber(Application.Contents("_STAT-PI-delay")/Application.Contents("_STAT-PI-count"),2)
If Me.Show Then
w "/*--- END DEBUGGER -----"
w "FWX_QueryCost="&FWX_QueryCost
w "_STAT-PI-count="&Application.Contents("_STAT-PI-count")
w "_STAT-PI-delay="&Application.Contents("_STAT-PI-delay")
w "_STAT-PI-score="&Application.Contents("_STAT-PI-score")
w "Log entries="&Me.Count
If IsObject(Me.Warnings)Then
w "Number of Warnings="&Me.Warnings.Length
If Me.Warnings.Length>0 Then
w "!!!!!!!!!!!!!! WARNINGS !!!!!!!!!!!!"
w "<div style='color:red;'>"&Me.Warnings.JoinString("<hr/>")&"</div>"
w "!!!!!!!!!!!!!! /WARNINGS !!!!!!!!!!!!"
End If
End If
w "----- END DEBUGGER ---*/"
Else
If IsTest Then
If IsFront And Query.EXT="" Then
If FWX_ResponseType<>"html" Or IsTrue(Me.Silent)Or IsTrue(j("no-debug"))Then
Else
RW "<script>"
RW "try{"
RW "window.FWX_QueryCost='"&Left(FormatNumber(FWX_QueryCost,3),6)&"';"
RW "$('.FWX_QueryCost').text(window.FWX_QueryCost);"
RW "}catch(e){};"
RW "</script>"
End If
End If
End If
End If
End Sub
Public Default Sub Debug(ByVal m)
If My_DBG=True Then Exit Sub
If Not(Me.Active Or Me.Show)Then Exit Sub
Me.Count=Me.Count+1
Me.LastStep=Timer()-Me.ThisStep
Me.Ellapsed=Timer()-Me.Start
Me.ThisStep=Timer()
Dim n,e,t,c
n=Digits(Me.Count,5)
e=Left(FormatNumber(Me.Ellapsed,6,0),6)
t=Left(FormatNumber(Me.LastStep,8,0),6)
c=IIf(t>1.000,"#f00",IIf(t>0.100,"#900",IIf(t>0.010,"#00C",IIf(t>0.001,"#777","#AAA"))))
s=Replace(Replace(Me.WaitingToPrint,"ccccc",c),"ttttt",t)
If m<>"" Then
Me.WaitingToPrint="<div class='pHide' style='color:ccccc'>"&n&"["&e&"](ttttt): "&m&"</div>"
Else
Me.WaitingToPrint=""
End If
If Me.Show Then
If Me.Filter="" Or InStr(1,m,Me.Filter)>0 Then
Response.Write s
Response.Flush()
End If
End If
End Sub
Public Function SaveLog()
SaveLogTo MyDomainCache&"History/"&RemoveSymbols(Now())&".html"
End Function
Public Function SaveLogTo(ByVal sFile)
WriteToFile sFile,FullLog
End Function
Public Function FullLog()
FullLog=Me.History.GetString
End Function
Function Stamp()
Stamp=Replace(Time()&":"&Right(FormatNumber(Timer(),2),2),":","")&"-"
End Function
Public TimerStart
Function StartTime()
TimerStart=Timer()
End Function
Public LastLap
Function Lap()
prev=LastLap
LastLap=Timer()
If prev>0 Then
Lap=Timer()-prev
Exit Function
End If
Lap=Timer()-TimerStart
End Function
Function Since()
Since=Timer()-TimerStart
TimerStart=0
End Function
Function EmailError(ByVal subj,ByVal body)
DBG "DEBUGGER - Emailerror - "&subj
SendEmail "error@"&MyDomain&"","dev@fyneworks.com",subj,body
DBG "DEBUGGER - Emailerror - DONE"
End Function
Function EmailStatus(ByVal subj,ByVal body)
DBG "DEBUGGER - EmailStatus - "&subj
SendEmail "status@"&MyDomain&"","dev@fyneworks.com",subj,body
DBG "DEBUGGER - EmailStatus - DONE"
End Function
Function RaygunLogError(ByVal subj,ByVal body)
End Function
Function RaygunLogStatus(ByVal subj,ByVal body)
End Function
Function LogError(ByVal subj,ByVal body)
LogError=Me.EmailError(subj,body)
End Function
Function LogEvent(ByVal subj,ByVal body)
LogEvent=Me.EmailStatus(subj,body)
End Function
Function Status(ByVal s)
DBG "DEBUGGER - Status - "&s
Dim subj,body
subj="P1!"&subj&Trim(rxReplace(Left(s,100),"[^\w\s\-\:\,\.\;\/\!\(\)\@\_]+"," "))&"-"&Me.Stamp&""
body=s&"<br/><br/>"&Me.Snapshot&"<br/><br/>"
Status=Me.LogEvent(subj,body)
DBG "DEBUGGER - Status - DONE"
End Function
Function Warning(ByVal s)
DBG "DEBUGGER - Warning - "&s
Dim subj,body
If IsObject(Me.Warnings)Then
Me.Warnings.Add "<div class='warning' id='warning-"&(Me.Warnings.Length+1)&"'>"&Encode(s)&"</div>"
End If
subj="P2!"&subj&Trim(rxReplace(Left(s,100),"[^\w\s\-\:\,\.\;\/\!\(\)\@\_]+"," "))&"-"&Me.Stamp&""
body=s&"<br/><br/>"&Me.Snapshot&"<br/><br/>"
Warning=Me.LogEvent(subj,body)
DBG "DEBUGGER - Warning - DONE"
End Function
Function Alert(ByVal s)
DBG "DEBUGGER - Alert - "&s
Dim subj,body
subj="P3!"&subj&Trim(rxReplace(Left(s,100),"[^\w\s\-\:\,\.\;\/\!\(\)\@\_]+"," "))&"-"&Me.Stamp&""
body=s&"<br/><br/>"&Me.Snapshot&"<br/><br/>"
Alert=Me.LogError(subj,body)
DBG "DEBUGGER - Alert - DONE"
End Function
Function Error(ByVal s)
DBG "DEBUGGER - Error - "&s
Dim subj,body
subj="P4!"&subj&Trim(rxReplace(Left(s,100),"[^\w\s\-\:\,\.\;\/\!\(\)\@\_]+"," "))&"-"&Me.Stamp&""
body=s&"<br/><br/>"&Me.Snapshot&"<br/><br/>"
Error=Me.LogError(subj,body)
DBG "DEBUGGER - Error - DONE"
End Function
Function ServerInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Server</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
For Each f In Request.ServerVariables
v=Request.ServerVariables(f)
o.Add "<div>"&f&": "&Encode(IIf(Len(v)>256,Len(v)&"b",v))&"</div>"
Next
o.Add "</div>"
ServerInfo=o
End Function
Function ApplicationInfo()
Exit Function
End Function
Function SessionInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Session Info</strong>"
If IsObject(Session)Then
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
For Each f In Session.Contents
v=Session.Contents(f)
o.Add "<div>"&f&": "&Encode(IIf(Len(v)>256,Len(v)&"b",v))&"</div>"
Next
o.Add "</div>"
ElseIf IsObject(MySession)Then
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
For Each f In MySession.Contents
v=MySession(f)
o.Add "<div>"&f&": "&Encode(IIf(Len(v)>256,Len(v)&"b",v))&"</div>"
Next
o.Add "</div>"
Else
o.Add "<div>Not available - Session object was not available</div>"
End If
SessionInfo=o
End Function
Function RuntimeInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Runtime</strong>"
If IsObject(j)Then
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
o.Add "FILES="&Replace(j("_STAT-FSO.ReadFile-List"),"|||","<br/>")
o.Add "<u style='padding:15px 0 0 0; display:block;'>Disk Activity</u>"
o.Add "<div>Checks: <strong>"&j("_STAT-FSO.FileExists")&" files</strong>, "&j("_STAT-FSO.FolderExists")&" folders</div>"
o.Add "<div>Opens: <strong>"&j("_STAT-FSO.GetFile")&" seeks</strong>, "&j("_STAT-FSO.ReadFile")&" reads</div>"
o.Add "<div>Templates: "&j("_STAT-TMPL-F")&" from file, "&j("_STAT-TMPL-M")&" from memory</div>"
o.Add "<u style='padding:15px 0 0 0; display:block;'>Template Parser Stats</u>"
o.Add "<div>MergeTemplate: "&j("_STAT-MergeTemplate")&", PreParse: "&j("_STAT-PreParse")&", Parse: "&j("_STAT-Parse")&"</div>"
o.Add "<u style='padding:15px 0 0 0; display:block;'>Performance</u>"
o.Add "<div>Boot Age: <strong>"&j("BOOT-AGE")&"</strong> minutes</div>"
o.Add "<div>Performance Stats: <strong>"&Application.Contents("_STAT-PI-count")&"</strong> requests in "&Application.Contents("_STAT-PI-delay")&"s</div>"
o.Add "<div>Performance Index: <strong>"&Application.Contents("_STAT-PI-score")&"</strong></div>"
o.Add "<u style='padding:15px 0 0 0; display:block;'>Database Calls: "&VOZ(j("_STAT-OpenRS"))&"</u>"
Dim s
For Each s In Split(j("_STAT-OpenRS-List"),vbCRLf)'"/*-*/")
s=Replace(s,"[/**/","[")
s=rxReplace(s,"(\s*(WHERE|FROM|ORDER|GROUP|HAVING|(LEFT|RIGHT|INNER)\s+JOIN))\b","<br/>          "&vbCR&vbLf&"$1")
o.Add "<div style=""font-family:'Courier New', Courier, monospace; border-bottom:#e7e7e7 solid 1px; padding:3px 0;"">"&vbCRLf&s&vbCRLf&"</div>"
Next
o.Add "</div>"
Else
o.Add "<div>Not available - j object was terminated</div>"
End If
RuntimeInfo=o
End Function
Function FyneworksInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Fyneworks Info</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
o.Add "<div>FWX_VERSION: "&FWX_VERSION&"</div>"
If IsObject(FWX)Then
If FWX.DataSource<>"" And(IsLocal Or IsAdmin)Then
o.Add "DataSource: "&rxReplace(FWX.DataSource,"^([^;]+);([^;]+)","$1 / $2")&Len(FWX.DataSource)&" bytes"
End If
o.Add "ConnectionString: "&Len(FWX.ConnectionString)&" bytes"
If(IsLocal Or IsAdmin)Then
o.Add "<div>ConnectionString (local/admin only): "&(FWX.ConnectionString)&"</div>"
End If
For Each f In Split("domain,site,captcha,gzip,CDNs",",")
If f<>"" Then o.Add "<div>FWX.Setting("""&f&"""): "&FWX.Setting(f)&"</div>"
Next
If(IsLocal Or IsAdmin)Then
For Each f In Split("ConnectionString,-key",",")
If f<>"" Then o.Add "<div>FWX.Setting("""&f&"""): "&FWX.Setting(f)&"</div>"
Next
End If
Else
o.Add "<div>FWX object was not available</div>"
End If
o.Add "<div>TMPL_FOLDER: "&TMPL_FOLDER()&"</div>"
o.Add "<div>SITE_FOLDER: "&SITE_FOLDER()&"</div>"
o.Add "<div>SETTINGS.asp: "&SITE_FOLDER()&"SETTINGS.asp</div>"
For Each f In Split("PAGE,PAGEID,PATH,PATHID,FULLID,FULLREQUEST,SCRIPT,URL,RAWURL,IsLocal,IsLive,IsFront,IsBack,IsHuman,IsRobot,Is404,IsHTTPReq,IsAjax,IsFF,IsIE,BrowserCode,IsMobile,IsDesktop,IsChrome,IsSafari,IsOpera,MyQuote,MyQuoteMerchant,MyAccount,MyMerchant,"&IIf(IsObject(Query),"Query.QString,Query.FString,","")&"MySite,MyDomain,MyWWWDomain,DOMAIN_NAME,SERVER_NAME,IsSubDomain,IsBOSub,SUBDOMAIN,",",")
If f<>"" Then
v=Eval(f)
o.Add "<div>"&f&": "&IIf(Len(v)>256,Len(v)&"b",v)&"</div>"
End If
Next
o.Add "</div>"
FyneworksInfo=o
End Function
Function ClientInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Client</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
Dim f
o.Add "<div>REQUEST: <a target='_blank' href='http://"&SERVER_NAME&FULLREQUEST&"'>see it</a> - "&SERVER_NAME&FULLREQUEST&"</div>"
For Each f In Split("MyName,MyEmail,",",")
If f<>"" Then
o.Add "<div>"&f&": <a href='http://"&SERVER_NAME&"/x/search?q="&Eval(f)&"'>"&Eval(f)&"</a></div>"
End If
Next
For Each f In Split("UAID,OSID,",",")
If f<>"" Then
v=Eval(f)
o.Add "<div>"&f&": "&IIf(Len(v)>256,Len(v)&"b",v)&"</div>"
End If
Next
o.Add "<div>MyVisitor: <a target='_blank' href='http://"&SERVER_NAME&"/x/@live.visitors@/f?q="&MyVisitor&"'>"&MyVisitor&"</a></div>"
o.Add "<div>VISITOR_IP: <a target='_blank' href='http://www.geoiptool.com/en/?IP="&VISITOR_IP&"'>"&VISITOR_IP&"</a></div>"
o.Add "<div>CF-Connecting-IP: "&Request.ServerVariables("HTTP_X-CF-Connecting-IP")&"</div>"
o.Add "<div>X-Forwarded-For: "&Request.ServerVariables("HTTP_X-Forwarded-For")&"</div>"
o.Add "</div>"
ClientInfo=o
End Function
Function FormInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Form</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
If IsSOAP Then
o.Add "SOAP DATA="
Else
If Request.Form<>"" Then
For Each f In Request.Form
v=Encode(Request.Form(f))
o.Add "<div>"&f&": "&Encode(IIf(Len(v)>256,Len(v)&"b",v))&"</div>"
Next
End If
End If
o.Add "</div>"
FormInfo=o
End Function
Function QueryInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
If Request.QueryString<>"" Then
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Query</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
For Each f In Request.QueryString
If f<>"" Then o.Add f&": "&Encode(Request.QueryString(f))&"<br/>"
Next
o.Add "</div>"
End If
QueryInfo=o
End Function
Function LastErrorInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
Set objError=Server.GetLastError()
If IsObject(objError)And objError.Line>0 Then
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Server.GetLastError</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
With objError
o.Add "ASPCode: "&.ASPCode()&"<br/>"
o.Add "ASPDescription: "&TOT(.ASPDescription(),.Description)&"<br/>"
o.Add "File: "&.Me.File()&"<br/>"
o.Add "Line: "&.Line()&"<br/>"
End With
o.Add "</div>"
End If
LastErrorInfo=o
End Function
Function InvoiceInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
If GTZ(j("invoice"))Then
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Invoice "&j("invoice")&"</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
o.Add("http://"&SERVER_NAME&"/x/invoices/"&j("invoice")&"")
o.Add "</div>"
End If
InvoiceInfo=o
End Function
Function QuoteInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
If GTZ(j("quote"))Then
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Quote "&j("quote")&"</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
o.Add("http://"&SERVER_NAME&"/x/quotes/"&j("quote")&"")
o.Add "</div>"
End If
QuoteInfo=o
End Function
Function AccountInfo()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
Dim f
If GTZ(j("account"))Then
o.Add "<strong style='padding:15px 0 0 0; display:block; font-size:110%;'>Account "&j("account")&"</strong>"
o.Add "<div class='debug-query-section' style='margin:0 0 0 1px; clear:both;'>"
o.Add("http://"&SERVER_NAME&"/x/accounts/"&j("account")&"")
o.Add "</div>"
End If
AccountInfo=o
End Function
Function MiniSnapshot()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
o.Add Me.RuntimeInfo
o.Add Me.ClientInfo
o.Add Me.SessionInfo
o.Add Me.QueryInfo
o.Add Me.History
MiniSnapshot=o
End Function
Function Snapshot()
Dim o:Set o=New QuickString:o.Divider=vbCRLf
o.Add Me.ClientInfo
o.Add Me.InvoiceInfo
o.Add Me.QuoteInfo
o.Add Me.AccountInfo
o.Add Me.SessionInfo
o.Add Me.RuntimeInfo
o.Add Me.FormInfo
o.Add Me.QueryInfo
o.Add Me.LastErrorInfo
o=Replace(o,"[","[ ")
Snapshot=o
End Function
Public Function Notify(ByVal s)
If IsTest And IsStaffPC Then
DBG "***** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
DBG "***** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
DBG "***** !!!! NOTIFICATION"
DBG s
If IsHuman And(IsTest Or IsAdmin Or IsStaffPC)Then
j("DEBUGGING")=j("DEBUGGING")&"<div class='DEBUGGING Hidden pHide' style='display:none;'>"&s&"</div>"
End If
DBG "***** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
DBG "***** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
End If
End Function
Public Function Announce(ByVal s)
Me.Debug "**************************************************************************"
Me.Debug s&""
Me.Debug "**************************************************************************"
End Function
Public Function HTTP404(ByVal s)
Me.HTTPError 404,s
End Function
Public Function HTTP403(ByVal s)
Me.HTTPError 403,s
End Function
Public Function HTTP500(ByVal s)
If IsGoogle Then
Me.HTTPError 503,s
Else
Me.HTTPError 500,s
End If
End Function
Public Function FatalError(ByVal s):Me.HTTP500 s:End Function
Public Function HTTPError(status,s)
Dim m:m=s
If 1=0 Or rxTest(s,"was deadlocked on lock resources with another process")Or rxTest(s,"query timeout expired")Or rxTest(s,"deadlock victim")Then
Dim bRetry,iRetry,sRetry
bRetry=False
iRetry=VOZ(Query.FQ("TimeoutRetryAttempt"))
sRetry=rxReplace(rxReplace(Query.FULLURL,"TimeoutRetryAttempt=[^&]+",""),"(\&|\?)+$","")
sRetry=sRetry&Iif(rxTest(sRetry,"\?"),"&","?")&"TimeoutRetryAttempt="&iRetry+1
If IsPost Then
ElseIf rxTest(SERVER_NAME,"epos|app")Then
ElseIf iRetry>=3 Then
DBG.Warning "Timeout cannot be avoided (tried 3 times)"
ElseIf IsRobot Then
ElseIf IsFront Then
bRetry=True
ElseIf IsBack Then
bRetry=True
End If
If bRetry Then
DBG.Warning "Timeout avoided, retrying... (attempt #"&iRetry&")"
Redirect sRetry
Response.End()
End If
m="Took too long (timeout). Please try again."
End If
Me.Announce "HTTP"&status&" Error: "&s
If IsObject(Query)Then If Query.FQ("noErrors")<>"" Then DIE s
If status>=500 Then
If IsObject(Query)And IsObject(MyCookies)Then
Me.Alert "FATAL: "&s
Else
Me.Announce "Cannot report error - environment wasn't ready..."
End If
End If
If IsHTTPReq Then
FWX_ResponseStatus 200
s="<a class='SendRequest-Error-A auto' href='"&Query.RAWURL&"' title='DBG.HTTPError: "&Encode(s)&"' style='background:#f00;color:#fff;padding:2px;' onclick='AjaxReplace(this);return false;' ondblclick='window.confirm('Server-side request to this URL failed',this.title)'>HTTP"&Status&" "&m&" (<u>try again</u>)</a>"
Else
FWX_ResponseStatus status
End If
If rxTest(EXT,"js(on)?")Or IsAjax Then
RC
RTXT
RW "{"&vbCRLf
RW """fwx"":""http error "&status&""","&vbCRLf
RW """error"":"""&EncodeJSON(s)&""","&vbCRLf
RW """status"":"""&status&""","&vbCRLf
RW """message"":"""&EncodeJSON(m)&""","&vbCRLf
For Each thing In Split("IsHTTPReq,IsAjax,IsPassive,MyAccount,MyEmail,MyCredentials,MyState,MyVisitor,MyQuote,MyVIP,MyCorporate",",")
RW """"&LCase(rxReplace(thing,"^(My|Is)",""))&""":"""&EncodeJSON(EXE(thing))&""","&vbCRLf
Next
If IsObject(visitor)Then
For Each thing In Split("visitor.VisitorOndemand,visitor.VisitorStateLoaded",",")
RW """"&Replace(thing,".","_")&""":"""&EncodeJSON(EXE(thing))&""","&vbCRLf
Next
End If
RW """uri"":"""&EncodeJSON(FULLREQUEST)&""","&vbCRLf
If IsTest Or IsAdmin Then
RW """sql"":"""&Replace(j("_STAT-LastRS"),"""","")&""","&vbCRLf
End If
If status=403 Then
RW """login"":""true"","&vbCRLf
If IsBack Then
RW """js"":""fwx.dialog('/x/auth/login',{height:250,width:300,id:'bo-ajax-login'})"","&vbCRLf
End If
End If
If IsTest Then
RW """last_request"":"""&Replace(j("__last_request"),"""","")&""","&vbCRLf
RW """last_response"":"""&Len(j("last_response"))&""","&vbCRLf
End If
RW """date"":"""&Now()&""""&vbCRLf
RW "}"
ElseIf rxTest(EXT,"jpe?g|png|gif|bmp|tiff?")Then
RTXT
RW s&""
ElseIf rxTest(EXT,"txt|csv|css")Then
RTXT
RW s&""
ElseIf rxTest(EXT,"xml")Then
RXML
RW "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"&vbCRLf
RW "<xml><response>"&vbCRLf
RW "<error>"&XmlEncode(s)&"</error>"&vbCRLf
RW "<message>"&XmlEncode(m)&"</message>"&vbCRLf
RW "<request>"&XmlEncode(FULLREQUEST)&"</request>"&vbCRLf
RW "</response></xml>"&vbCRLf
ElseIf IsHTTPReq Then
RHTM
RW s&""
Else
If IsBack Or IsFront Or IsStaffPC Or rxTest(SERVER_NAME,"^("&FWX__BO_Reserved_SubDomains&")")Then
Else
If FileExists("/maintenance.asp")Then
Transfer "/maintenance.asp"
End If
End If
RHTM
RW "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"
RW "<html xmlns='http://www.w3.org/1999/xhtml'>"
RW "<head>"
RW "<meta http-equiv='Content-Type' content='text/html; charset=ISO-8859-1' />"
RW "<style title='text/css'>"
RW "*{ cursor:default; margin:0; padding:0; font-size:100%; }"
RW "body{ margin:10px; background:#fff; color:#555; font-family:Arial, Helvetica, sans-serif; font-size:80%; }"
RW "h1,h3,p{ margin:5px; }"
RW "#xxx{ }"
RW "#err{ margin:30px 0 0 0; background:#eee; border:#aaa solid 0px; "&IIf(IsTest,"","display:none;")&"}"
RW "#err legend{ text-decoration:underline; background:#ddd; color:#999; }"
RW "#err p{ font-size:80%; }"
RW "#det{ font-size:80%; }"
RW "#env{ padding:20px; }"
RW "</style>"
RW "<scr"&"ipt src='//code.jquery.com/jquery-latest.pack.js' type='text/javascript'></scr"&"ipt>"
RW "</head>"
RW "<body ondblclick=""$('#err').show();"">"
Dim e
If e<>"" Then
RW e
Else
RW "<div id='msg'>"
RW "<div>"
RW "<strong>Oops!</strong> We cannot show you this page right now (X3)."
RW "</div>"
RW "<div>"
RW "Please <strong>try again soon</strong>, <a href='/'>visit our homepage</a> "
RW "or <a href='javascript:history.go(-1);'>click here</a> to go back."
RW "</div>"
RW "</div>"
End If
RW "<fieldset id='err'>"
RW "<legend id='det'>"
RW "Error Information"
RW "</legend>"
RW "<div id='env'>"
RW "<strong>Description</strong>"
RW "<div id='xxx'>"&s&"</div>"
If Err.Description<>"" Or Err.Number<>0 Then
RW "<div>Err.Number: "&Err.Number&"</div>"
RW "<div>Err.Description: "&Err.Description&"</div>"
End If
If IsObject(j)Then
If IsSet(ColItem(j,"FatalError-Dump"))Then
RW "<div>Err.Detail: "&ColItem(j,"FatalError-Dump")&"</div>"
End If
End If
If IsAdmin Or IsLocal Then
RW "<br/>"
RW "<h2>Environment Snapshot</h2>"
If IsObject(Query)Then
RW Me.Snapshot()
Else
RW "Snapshot isn't available because environment wasn't ready, but here's some stuff:"
RW "<br/><br/><h2>Application</h2>"&PrintCol(Application.Contents)
End If
End If
RW "</div>"
RW "</fieldset>"
RW "</body>"
RW "</html>"
End If
RE
End Function
Public Silent
Public Function Silence()
Me.Silent=True
End Function
function dump(byRef mixed)
select case lcase(typename(mixed))
case "string":
dump="("&typename(mixed)&") """&mixed&""""
case "iregexp2":
dump="(RegExp) /"&mixed.pattern&"/"&iif(mixed.global,"g","")&iif(mixed.ignoreCase,"i","")&iif(mixed.multiline,"m","")
case "variant()":
dump=dump_hdlArray(mixed)
case "dictionary":
dump=dump_hdlDictionary(mixed)
case "recordset":
dump=dump_hdlRecordset(mixed)
case "stream":
dump=dump_hdlStream(mixed)
case else:
on error resume next
dump="("&typename(mixed)&") "&mixed
if(Err.number<>0)then
dump="("&typename(mixed)&")"
Err.clear()
end if
on error goto 0
end select
end function
function dump_hdlArray(byRef arr)
dim entry:dump_hdlArray=""
for each entry in arr
dump_hdlArray=dump_hdlArray&","&dump(entry)
next
dump_hdlArray="(Array) ["&mid(dump_hdlArray,2)&"]"
end function
function dump_hdlDictionary(byRef Dic)
dim key:dump_hdlDictionary=""
for each key in Dic.keys()
dump_hdlDictionary=dump_hdlDictionary&","""&key&""": "&dump(Dic(key))
next
dump_hdlDictionary="(Dictionary) {"&mid(dump_hdlDictionary,2)&"}"
end function
function dump_hdlRecordset(byRef Rs)
dim i,row:dump_hdlRecordset=""
while(not Rs.eof)
row=""
for i=0 to Rs.Fields.count-1
ON ERROR RESUME NEXT:Err.clear()
row=row&","""&Rs(i).name&""": "&dump(Rs(i).value)
if(Err.number<>0)then
row=row&","""&Rs(i).name&""": "&dump(ccur(Rs(i).value))
end if
ON ERROR GOTO 0
next
dump_hdlRecordset=dump_hdlRecordset&",(Record) {"&mid(row,2)&"}"
Rs.moveNext()
wend
dump_hdlRecordset="(Recorset) ["&mid(dump_hdlRecordset,2)&"]"
end function
function dump_hdlStream(byRef St)
if(St.type=1)then
dump_hdlStream="(Stream) adTypeBinary"
else
St.position=0
dump_hdlStream="(Stream) """&St.readText()&""""
end if
end function
End Class
Class fwxValidation
Public lclErrors
Public Property Let Errors(str)
If IsEmpty(str)Then Exit Property
If IsSet(lclErrors)Then lclErrors=lclErrors&"~"
lclErrors=lclErrors&str
End Property
Public Property Get Errors()
Errors=lclErrors
End Property
Function ClearErrors()
lclErrors=Empty
End Function
Public lclRules
Public Property Let Rules(str)
If IsEmpty(str)Then Exit Property
If IsSet(lclRules)Then lclRules=lclRules&","
lclRules=lclRules&str
End Property
Public Property Get Rules()
Rules=Split(lclRules,",")
End Property
Public Property Get RulesString()
RulesString=lclRules
End Property
Function ClearRules()
lclRules=Empty
End Function
Function HasErrors()
HasErrors=IsSet(lclErrors)
End Function
Function OK()
OK=IsEmpty(lclErrors)
End Function
Public Property Get ErrorFields()
Dim o:Set o=New QuickString:o.Divider=","
Dim e:For Each e In Split(lclErrors,"~")
o.Add rxSelect(e,"^\w+")
Next
ErrorFields=o
End Property
Public Property Get AllErrors()
AllErrors=Me.ErrorsText(Me.ErrorFields)
End Property
Public Property Get Field(fld)
r.Pattern="\W*"&fld&"\W*"
Field=r.Test(lclErrors)
End Property
Public Property Get ListErrors(fld):ListErrors=ErrorsHTML(fld):End Property
Public Property Get ErrorsHTML(fld):ErrorsHTML=DoListErrors(fld,True):End Property
Public Property Get ErrorsText(fld):ErrorsText=DoListErrors(fld,False):End Property
Public Property Get DoListErrors(ByVal fld,ByVal bHTML)
output=""
r.Pattern="\W*"&fld&"\W*\:[^\~]*"
Set valErrors=r.Execute(lclErrors)
For Each e In valErrors
e=Replace(e,",","")
e=Mid(e,InStr(1,e,":")+1)
If InStr(1,e,"-")>0 Then
tmp=Split(e,"-")
e=Trim(tmp(0))
n=Trim(tmp(1))
End If
Select Case LCase(e)
Case "captcha","recaptcha","antispam","honey":
DBG.Warning "Anti-Spam Failed"
d="failed anti-spam check"
Case "empty","missing":d="required"
Case "short":d="too short, "&VOV(n,50)&" chars min."
Case "long":d="too long, "&VOV(n,1024)&" chars max."
Case "small":d="too small, must be at least "&n
Case "big":d="too big, must be less than "&n
Case "gt":d="must be less than "&n&""
Case "lt":d="must be more than "&n&""
Case "eq":d="must not be '"&Slice(n,15)&"'"
Case "df":d="must be equal to '"&Slice(n,15)&"'"
Case "neq":d="must not be '"&VOZ(n)&"'"
Case "ndf":d="must be equal to '"&VOZ(n)&"'"
Case "number","nan":d="not a valid number"
Case "email":d="invalid email address"
Case "postcode":d="invalid postcode"
Case "phone":d="invalid phone #"
Case "url":d="invalid url (start with 'http://...')"
Case "id":d="must not contain any symbols or spaces"
Case "match","mismatch":d=IIf(IsSet(n),"does not match '"&n&"'","fields do not match")
Case Else:d=Replace(TOT(e,"ERROR"),"_"," ")
End Select
If bHTML Then
output=output&d&""'&"<img src="""&App.Image("menu/error")&""" hspace='0' vspace='0' border='0' height='15' width='15'/>"& hsp(5)&d&""
Else
output=output&d&""
End If
Next
Set valErrors=Nothing
If IsEmpty(output)Then Exit Property
DoListErrors=output
End Property
Public Function Decode(ByVal str)
r.Pattern="\[\w+\]"
Set matches=r.Execute(str)
For Each m In matches
str=Replace(str,m,col(Mid(m,2,Len(m)-1)))
Next
End Function
Public Function Validate()
Dim strErrors:strErrors=""
Dim fldErrors:fldErrors=""
Dim aFields:aFields=Me.Rules
For Each sField In aFields
sField=Trim(sField)
If IsSet(sField)Then
aRules=Split(sField,":")
fld=aRules(0)
If 1=0 Then
ElseIf rxTest(LCase(fld),"(recaptcha|captcha)")Then
If Not reCAPTCHA_Check()Then
strErrors=strErrors&"recaptcha:recaptcha~"
strErrors=strErrors&"recaptcha_response_field:recaptcha~"
End If
ElseIf rxTest(LCase(fld),"(honey|spam|antispam)")Then
j("recaptcha-check")="y"
If Query.FQ("honey,honey_email,honey_name,honey_url")="" AND col("honey")="" AND col("honey_email")="" AND col("honey_name")="" AND col("honey_url")="" Then
Else
j("recaptcha-check")="n"
strErrors=strErrors&"honey:honey~"
End If
Else
SkipMe=False
fldError=Empty
iRuleCount=1
Do While iRuleCount<=UBound(aRules)And fldError=Empty And SkipMe=False
If InStr(1,aRules(iRuleCount),"-")>0 Then
rs=Split(Trim(aRules(iRuleCount)),"-")
rule=rs(0)
Else
rule=Trim(aRules(iRuleCount))
End If
rule=LCase(rule)
Select Case rule
Case "skip","notrequired","notreq","notrq","allownull":
If IsEmpty(col(fld))Then SkipMe=True
Case "required":
If Not IsSet(col(fld))Then fldError=fld&":missing"
Case "email":
If Not IsEmail(Trim(col(fld)))Then fldError=fld&":email"
Case "number":
If Not IsNumeric(Trim(col(fld)))Then fldError=fld&":number"
Case "postcode":
If Not IsPostcode(Trim(col(fld)))Then fldError=fld&":postcode"
Case "phone":
If Not IsPhone(Trim(col(fld)))Then fldError=fld&":phone"
Case "url":
If Not IsUrl(Trim(col(fld)))Then fldError=fld&":url"
Case "id":
If Not IsValidID(Trim(col(fld)))Then fldError=fld&":id"
Case Else
Select Case LCase(rule)
Case "match"
If Trim(Trim(col(fld)))<>col(rs(1))Then fldError=fld&":mismatch-"&rs(1)
Case "maxlen"
If Len(Trim(col(fld)))>VOZ(rs(1))Then fldError=fld&":long-"&rs(1)
Case "minlen"
If Len(Trim(col(fld)))<VOZ(rs(1))Then fldError=fld&":short-"&rs(1)
Case "max"
If VOZ(Trim(col(fld)))>VOZ(rs(1))Then fldError=fld&":big-"&rs(1)
Case "min"
If VOZ(Trim(col(fld)))<VOZ(rs(1))Then fldError=fld&":small-"&rs(1)
Case "df"
If Trim(col(fld))=Trim(rs(1))Then fldError=fld&":eq-"&rs(1)
Case "eq"
If Trim(col(fld))<>Trim(rs(1))Then fldError=fld&":df-"&rs(1)
Case "ndf"
If VOZ(col(fld))=VOZ(rs(1))Then fldError=fld&":neq-"&rs(1)
Case "neq"
If VOZ(col(fld))<>VOZ(rs(1))Then fldError=fld&":ndf-"&rs(1)
End Select
End Select
If SkipMe Then
strErrors=""
ElseIf fldError<>"" Then
strErrors=strErrors&fldError&"~"
End If
iRuleCount=iRuleCount+1
Loop
End If
End If
Next
Me.Errors=strErrors
Validate=IsEmpty(strErrors)
End Function
Public col
Private Sub Class_Initialize()
Set col=Dictionary()
UpdateCollection col,Query.Fata
UpdateCollection col,Query.Data
End Sub
End Class
Dim global_Validation
Function validation()
If Not IsObject(global_Validation)Then Set global_Validation=New fwxValidation
Set Validation=global_Validation
End Function
Class Zebra
Private i
Private s
Public n
Public Stripes
Public Row
Public RowS
Public Sub Add
Me.Stripes=VOV(Me.Stripes,2)
i=i+1
s=s+1
n=n+1
If(s Mod 2)=0 Then
sClass="RowS2"
Else
sClass="RowS"
End If
Select Case Stripes
Case 2:
If(i Mod 2)=0 Then
rClass="RowA"
Else
rClass="RowB"
End If
Case 3:
If(i Mod 3)=0 Then
rClass="RowA"
i=0
ElseIf(i Mod 2)=0 Then
rClass="RowB"
Else
rClass="RowC"
End if
End Select
Me.Row=rClass
Me.RowS=sClass
End Sub
Public Function RowID()
RowID="ZebraRow"&Me.n
End Function
Public Function DoStripe(ByVal blnEffect,ByVal sClass)
Me.Add()
output=Empty
If blnEffect Then
output=output&Me.Effect&" "
End If
output=output&" id="""&Me.RowID&""" class="""&TOT(sClass,Me.Row)&""" "
DoStripe=output
End Function
Public Function StripeOnly()
StripeOnly=DoStripe(False,Empty)
End Function
Function Stripe()
Stripe=DoStripe(True,Empty)
End Function
Function StripeWith(ByVal sClass)
StripeWith=DoStripe(True,sClass)
End Function
Public Function Effect()
Effect=" onmouseover=""AddClass('RowH', this);"" onmouseout=""RemoveClass('RowH', this);"" "
End Function
Public Function SEffect()
SEffect=" onclick=""ToggleClasses('"&Me.Row&"', '"&Me.RowS&"', this);"" "
End Function
Public Function CheckEffect()
CheckEffect=" zA='"&Me.Row&"' zH='"&Me.RowS&"' zId='ZebraRow"&Me.n&"' "
End Function
Public Sub Reset()
i=0
End Sub
End Class
Dim global_z
Function z()
If Not IsObject(global_z)Then Set global_z=New Zebra
Set z=global_z
End Function
Const fwxSearchEncode_Space="XXXXX-space-XXXXX"
Const fwxSearchEncode_Quote="XXXXX-squot-XXXXX"
Function fwxSearchEncode(ByVal s)
s=rxReplace(s,"'",fwxSearchEncode_Quote)
Dim m,x
For Each m In rxMatches(s,"""[^""]+""")
x=Replace(Replace(m," ",fwxSearchEncode_Space),"""","")
s=Replace(s,m,x)
Next
m="":x=""
fwxSearchEncode=s
End Function
Function fwxSearchDecode(ByVal s)
s=rxReplace(s,fwxSearchEncode_Space," ")
s=rxReplace(s,fwxSearchEncode_Quote,"'")
fwxSearchDecode=s
End Function
Function fwxSearchDecodeSafe(ByVal s)
s=rxReplace(s,fwxSearchEncode_Space," ")
s=rxReplace(s,fwxSearchEncode_Quote,"''")
fwxSearchDecodeSafe=s
End Function
Class fwxSearchField
Private lclName
Private lclTable
Private lclCompare
Private lclMatch
Private lclCriteria
Public Source
Public Data
Public Compare
Public Invert
Public Table
Public Match
Public Criteria
Public AltCriteria
Public DateFilter
Public Active
Private Sub Class_Initialize()
lclName="field"
lclMatch="o"
Me.DateFilter=TOT(Query.FQ("datefilter"),"#")
Me.Active=True
End Sub
Public Property Get Name()
str=lclName
If Me.Table<>Empty Then str=Me.Table&"."&str
Name=str
End Property
Public Property Let Name(str)
lclName=SQLSafeField(str)
End Property
Public Function SQL()
If Not Me.Active Then Exit Function
c=TOT(Me.Criteria,Me.AltCriteria)
c=fwxSearchEncode(c)
If Me.Data="boolean" Then
If LCase(c)<>"-empty" And c<>"-" Then
output="ISNULL("&Me.Name&","&IIf(IsTrue(TOT(Me.Altcriteria,"y")),1,0)&")="&Abs(CInt(IsTrue(c)))&""
Else
End If
ElseIf Me.Compare="set" Or Me.Compare="notempty" Then
Select Case Me.Data
Case "number":output="ISNULL("&Me.Name&",0) > 0"
Case "ntext":output="ISNULL("&Me.Name&",'') NOT LIKE ''"
Case Else:output="ISNULL("&Me.Name&",'') <> ''"
End Select
ElseIf Me.Compare="empty" Or Me.Compare="notset" Then
Select Case Me.Data
Case "number":output="ISNULL("&Me.Name&",0)=0"
Case "ntext":output="ISNULL("&Me.Name&", '') LIKE ''"
Case Else:output="ISNULL("&Me.Name&", '')=''"
End Select
ElseIf c<>"" Then
If 1=0 Then
ElseIf InStr(1,LCase(Me.Compare),"%q%")>0 Then
output=Me.Name&" "&Replace(Me.Compare,"%q%",c)
ElseIf LCase(c)="-empty" Or LCase(c)="!empty" Or c="-" Then
If Me.Data="ntext" Then
output=Me.Name&" LIKE '"&c&"'"
Else
output="ISNULL("&Me.Name&",'')=''"
End If
ElseIf Me.Compare="notempty" Or LCase(c)="-notempty" Or LCase(c)="!-empty" Then
output="ISNULL("&Me.Name&",'') <> ''"
ElseIf Me.Compare="gtz" Or LCase(c)="-gtz" Or LCase(c)="-gz" Or LCase(c)="!-zero" Then
output=Me.Name&">0"
ElseIf Me.Compare="ltz" Or LCase(c)="-ltz" Or LCase(c)="-lz" Then
output=Me.Name&"<0"
ElseIf Me.Compare="zero" Or LCase(c)="-isz" Or LCase(c)="-iz" Or LCase(c)="-zero" Or LCase(c)="-z" Then
output=Me.Name&"=0"
ElseIf Me.Compare="beginswith" Or Me.Compare="startswith" Or Me.Compare="startsin" Then
output=Me.Name&" LIKE '"&c&"%'"
ElseIf Me.Compare="endswith" Or Me.Compare="endsin" Then
output=Me.Name&" LIKE '%"&c&"'"
ElseIf Me.Compare="phrase" Or Me.Compare="like" Then
output=Me.Name&" LIKE '%"&c&"%'"
ElseIf Me.Compare="exactphrase" Then
If StartsWith(c,"-")Or StartsWith(c,"!")Then
output=Me.Name&"<>'"&Mid(c,2)&"'"
Else
output=Me.Name&"='"&c&"'"
End If
ElseIf Me.Compare="ilike" Then
output="'%"&c&"%' LIKE "&Me.Name&""
ElseIf Me.Compare="nlist" Then
If c<>"" Then
c_nlist=NumberList(c)
If c_nlist="0" And c<>"0" Then
ElseIf rxTest(c_nlist,"^[\d \,\-]+$")Then'c_nlist<>"0" Then
output=Me.Name&" IN (-8,"&c_nlist&")"
End If
End If
ElseIf Me.Compare="exact" Or Me.Compare="equals" Then
If StartsWith(c,"-")Or StartsWith(c,"!")Then
output=Me.Name&"<>'"&Mid(c,2)&"'"
Else
output=Me.Name&"='"&c&"'"
End If
ElseIf Me.Compare="perfect" Then
output=Me.Name&"='"&c&"'"
ElseIf Me.Compare="list" Then
If Me.Data="number" Then
If c<>"" Then
c_nlist=NumberList(c)
If c_nlist="0" And c<>"0" Then
ElseIf rxTest(c_nlist,"^[\d \,\-]+$")Then'c_nlist<>"0" Then
output=Me.Name&" IN (-9,"&c_nlist&")"'/*search.number.list*/"
End If
End If
Else
c=rxReplace(c,"[\r\n\t\s]*(,[\r\n\t\s]*)+",",")
c=rxReplace(c,",+",",")
c=rxReplace(c,"^,|,$","")
c=Trim(rxReplace(c,"[\r\n\t\s]+"," "))
If c="," Then c=""
If c<>"" Then
output=Me.Name&" IN ('"&Replace(Replace(c,",","','"),",'',",",")&"')"
Else
End If
End If
ElseIf Me.Compare="inlist" Then
c_nlist=rxReplace(Trim(c),"\s+",",")
If rxTest(c_nlist,"^[\d\,]+$")Then
output="("&Me.Name&"<>'' AND dbo.[InList]("&Me.Name&",'"&c_nlist&"')=1)"
End If
Else
output=""
MyWords=Split(c," ")
For Each word In MyWords
If lclName="q" Then
QueryParameters=Split(word,":")
If UBound(QueryParameters)>0 Then
If LCase(QueryParameters(0))<>LCase(lclName)Then
word=Empty
Else
word=Trim(QueryParameters(1))
End If
End If
End If
If Left(word,1)="!" Or Left(word,1)="-" Then
word=Mid(word,2)
InvertMe=True
ElseIf Left(word,1)="=" Then
word=Mid(word,2)
Me.Compare="exact"
Else
InvertMe=False
End If
Dim operator
operator=rxSelect(word,"^[\<\>\=]{1,2}")
If operator<>"" Then
word=Mid(word,Len(operator)+1)
Else
Me.Compare=TOT(Me.Compare,"equal"):operator=""
If InStr(1,Me.Compare,"more")>0 Then operator=">"
If InStr(1,Me.Compare,"less")>0 Then operator="<"
If InStr(1,Me.Compare,"equal")>0 Then operator=operator&"="
End If
If Trim(word)<>"" Then
Dim newstr:newstr=""
If 1=0 Then
ElseIf Me.Data="number" Then
word=SQLSafeNumber(word)
newstr=Me.Name&operator&VOZ(word)
ElseIf Me.Data="year" Then
word=SQLSafeNumber(word):If GTZ(word)Then newstr="DATEPART(yyyy,"&Me.Name&")='"&word&"'"
ElseIf Me.Data="yy" Then
word=SQLSafeNumber(word):If GTZ(word)Then newstr="DATEPART(yy,"&Me.Name&")='"&word&"'"
ElseIf Me.Data="month" Then
word=SQLSafeNumber(word):If GTZ(word)Then newstr="DATEPART(mm,"&Me.Name&")='"&word&"'"
ElseIf Me.Data="day" Then
word=SQLSafeNumber(word):If GTZ(word)Then newstr="DATEPART(dd,"&Me.Name&")='"&word&"'"
ElseIf Me.Data="weekday" Then
word=SQLSafeNumber(word):If GTZ(word)Then newstr="DATEPART(dw,"&Me.Name&")='"&word&"'"
ElseIf Me.Data="week" Then
word=SQLSafeNumber(word):If GTZ(word)Then newstr="DATEPART(wk,"&Me.Name&")='"&word&"'"
ElseIf Me.Data="date" Or rxTest(Me.Data,"^(localdate|absdate|dateonly)$")Then 'Or Me.Data="absdate" Or Me.Data="localdate" Or Me.Data="selecteddate" Then
If rxTest(word,"^([\d\w\/\-]*)$")Then
Else
word=rxReplace(word,"^([^\+\@\T]+)([\+\@\T])(\d+)","$1 $3")
End If
word=DateFromString(word)
If Not IsDate(word)Then
Else
Dim timezone,tzZulud,tzAlpha,tzOmega
timezone=MyTimezone
If rxTest(word,"\d{2}\:\d{2}")Then
word=FmtDate(word,"%Y-%M-%D %H:%N:%S")
Else
word=YYYY_MM_DD(word)
If InStr(1,operator,">")>0 Then word=word&" 00:00:00"
If InStr(1,operator,"<")>0 Then word=word&" 23:59:59"
End If
If Me.DateFilter<>"#" And operator="=" Then
If timezone="" Or timezone="-" Or rxTest(Me.Data,"^(localdate|absdate|dateonly)$")Then
newstr="CONVERT(nvarchar(8),"&Me.Name&",12)='"&yymmdd(word)&"'"
Else
word=YYYY_MM_DD(word)
tzAlpha=word&" 00:00:00"':00"
tzOmega=word&" 23:59:59"':59" MSSQL time rounding problem
tzAlphaZ=DateToUTC(tzAlpha,timezone)
tzOmegaZ=DateToUTC(tzOmega,timezone)
newstr=Me.Name&" BETWEEN '"&tzAlphaZ&"' AND '"&tzOmegaZ&"'"
End If
Else
If timezone="" Or timezone="-" Or rxTest(Me.Data,"^(localdate|absdate|dateonly)$")Then
Else
word=DateToUTC(word,timezone)
word=FmtDate(word,"%Y-%M-%D %H:%N:%S")
End If
newstr=Me.Name&operator&TOT(Me.DateFilter,"'")&word&TOT(Me.DateFilter,"'")&""
End If
End If
Else
If Me.Compare="instr" Then
newstr="INSTR("&Trim(field)&",'"&Trim(word)&"')>0"
ElseIf InStr(1,Me.Compare,"exact")>0 Or InStr(1,Me.Compare,"equals")>0 Then
newstr=Me.Name&"='"&word&"'"
ElseIf Me.Compare="not" Then
newstr=Me.Name&" <> '"&word&"'"
Else
If Not IsDate(Trim(word))Then
newstr=Me.Name&" LIKE '%"&word&"%'"
End If
End If
End If
If newstr<>Empty Then
If InvertMe=True Then newstr="NOT "&newstr&""
output=MergeCriteria(output,"AND",newstr)
End If
End If
Next
End If
End If
If output<>"" Then
If Me.Invert Then
output="NOT "&output
End If
output=fwxSearchDecodeSafe(output)
End If
SQL=output
End Function
End Class
Class fwxSearch
Private lclWithin
Private lclConfig
Public Config
Public Fields
Private Sub Class_Initialize()
Me.Reset()
End Sub
Public Sub Reset()
Set Fields=Nothing
Set Config=Nothing
Set Fields=Dictionary()
Set Config=Dictionary()
lclConfig=Empty
Me.Config("Active")=True
End Sub
Public Sub AutoConfig()
Me.Config("Criteria")=Query.FQ("q")
Me.Config("Fields")=TOT(SearchFields,Query.FQ("fields"))
Me.Within=(Query.FQ("within"))
Me.DefaultConfig=TOT(SearchConfig,Query.FQ("sconfig"))
Me.AddFilter Me.Config("Fields"),Me.Config("Criteria")
End Sub
Public Property Get DefaultConfig()
DefaultConfig=lclConfig
End Property
Public Property Let DefaultConfig(str)
If IsEmpty(str)Then Exit Property
lclConfig=Trim(str)
Parameters=Split(lclConfig,",")
For Each p in Parameters
Marker=InStr(1,p,"-")
Fieldr=Mid(p,Marker+1)
Select Case LCase(Left(p,Marker-1))
Case "c"
Me.Config("Compare")=LCase(Fieldr)
Case "d"
Me.Config("Data")=LCase(Fieldr)
Case "s"
Dim qi
qi=Replace(Replace(Fieldr,"+",","),"|",",")
If FieldTable<>"" Then qi=FieldTable&"."&Replace(Replace(FieldName,"[",""),"]","")&","&qi
Me.Config("AcceptedVars")=qi
Me.Config("Criteria")=Query.FQ(qi)
Case "m"
Me.Config("Match")=LCase(Fieldr)
Case "df"
Me.Config("DateFilter")=LCase(Fieldr)
Case "if"
If Fieldr="set" Then
xxx=Query.FQ(Fieldr&Enclose(Me.Config("AcceptedVars"),",",""))
Else
xxx=Query.FQ(Fieldr)
End If
Me.Config("Active")=CBool(xxx<>"" And xxx<>"-")
Case "nif"
xxx=Query.FQ(Fieldr)
Me.Config("Active")=CBool(Not(CBool(xxx<>"" And xxx<>"-")))
End Select
Next
End Property
Public Property Get Within()
Within=lclWithin
End Property
Public Property Let Within(str)
lclWithin=Trim(str)
End Property
Public Function AddFilter(FieldsToSearch,Expression)
If Expression="" Then
If IsObject(Query)Then
Expression=Query.FQ("q")
End If
End If
MyFields=Split(FieldsToSearch,",")
For Each field In MyFields
If Trim(field)<>Empty Then
FieldName=Empty
FieldTable=Empty
FieldParameters=Split(field,":")
FieldName=Trim(FieldParameters(0))
FieldName=TOT(Query.FQ(FieldName),FieldName)
If InStr(1,FieldName,".")>0 Then
FieldTable=Left(FieldName,InStr(1,FieldName,".")-1)
FieldName=Mid(FieldName,InStr(1,FieldName,".")+1)
End If
Set Me.Fields(field)=New fwxSearchField
With Me.Fields(field)
.Name=FieldName
.Table=TOT(FieldTable,Me.Config("Table"))
.Criteria=Me.Config("Criteria")
.Compare=Me.Config("Compare")
.Match=TOT(Me.Config("Match"),"o")
.DateFilter=TOT(Me.Config("DateFilter"),Query.FQ("datefilter"))
.Active=Me.Config("Active")
If UBound(FieldParameters)>0 Then
Dim iParameter
For iParameter=1 To UBound(FieldParameters)
p=FieldParameters(iParameter)
Marker=InStr(1,p,"-")
Fieldr=Mid(p,Marker+1)
If(Marker<=0)Or(Marker>0 And LCase(Left(p,2))="d-")Then
.Data=LCase(TOT(Fieldr,"text"))
Else
Select Case LCase(Left(p,Marker-1))
Case "c"
tmp_str=LCase(Fieldr)
If Left(tmp_str,1)="!" Then
.Invert=True
tmp_str=Mid(tmp_str,2,Len(tmp_str))
End If
.Compare=LCase(tmp_str)
Case "s"
Dim qi,qs
qi=Replace(Replace(Fieldr,"+",","),"|",",")
If FieldTable<>"" Then qi=FieldTable&"."&Replace(Replace(FieldName,"[",""),"]","")&","&qi
.Source=qi
qs=Query.FQ(qi)
.Criteria=qs
Case "m"
.Match=LCase(Fieldr)
Case "alt"
.AltCriteria=Query.FQ(Replace(Replace(Fieldr,"|",","),"+",","))
Case "def"
.AltCriteria=Fieldr
Case "if"
If Fieldr="set" Then
xxx=TOT(.Criteria,.AltCriteria)
Else
xxx=Query.FQ(Fieldr)
End If
.Active=CBool(xxx<>"" And xxx<>"-")
Case "nif"
xxx=Query.FQ(Fieldr)
.Active=CBool(Not(CBool(xxx<>"" And xxx<>"-")))
Case "j"
.Join=Fieldr
End Select
End If
Next
Else
End If
End With
End If
Next
End Function
Function SQL()
output=""
If Not IsObject(Me.Fields)Then Exit Function
Dim q,qs,dq
Set dq=Dictionary()
qs=Split(fwxSearchEncode(Query.FQ("q"))," ")
For Each q In qs
LetsInvert=False
If Left(q,1)="!" Or Left(q,1)="-" Then
q=Mid(q,2)
LetsInvert=True
End If
For Each field In Me.Fields
With Me.Fields(field)
If .Source="q" Then'Or .Source="" Then
.Criteria=q
Dim stmp
stmp=.SQL
dq(q)=MergeCriteria(dq(q),"OR",stmp)
End If
End With
Next
If dq(q)<>"" Then
If LetsInvert Then
dq(q)="NOT ("&dq(q)&")"
End If
output=MergeCriteria(Enclose(dq(q),"(",")"),"AND",output)
End If
Next
For Each field In Me.Fields
With Me.Fields(field)
Dim fsql
If .Source<>"q" Then
fsql=.SQL
If fsql<>"" Then
output=MergeCriteria(output,ANDOR(Me.Config("Match")),.SQL)
End If
End If
End With
Next
output=MergeCriteria(output,"AND",Me.Within)
SQL=output
End Function
End Class
Class XmlBuilder
Public Xml
Public Indents
Public Stack()
Private StackIndex
Public Sub Class_Initialize
StackIndex=0
Indents="  "
Xml="<?xml version=""1.0"" encoding=""UTF-8""?>"&VbCrLf
End Sub
Private Sub Indent()
Dim i,j
j=StackIndex
For i=0 To j
Xml=Xml&Indents
Next
End Sub
Public Sub Push(Elem,Attributes)
Indent()
Xml=Xml&"<"&Elem
If Attributes<>"" Then
Xml=Xml&" "&Attributes
End If
Xml=Xml&">"&VbCrLf
ReDim Preserve Stack(StackIndex)
Stack(StackIndex)=Elem
StackIndex=StackIndex+1
End Sub
Public Sub AddElement(Elem,Content,Attributes)
Indent()
Xml=Xml&"<"&Elem
If Attributes<>"" Then
Xml=Xml&" "&Attributes
End If
Xml=Xml&">"&Server.HTMLEncode(Content)&"</"&Elem&">"&VbCrLf
End Sub
Public Sub AddXmlElement(Elem,Content,Attributes)
Indent()
Xml=Xml&"<"&Elem
If Attributes<>"" Then
Xml=Xml&" "&Attributes
End If
Xml=Xml&">"&Content&"</"&Elem&">"&VbCrLf
End Sub
Public Sub EmptyElement(Elem,Attributes)
Indent()
Xml=Xml&"<"&Elem
If Attributes<>"" Then
Xml=Xml&" "&Attributes
End If
Xml=Xml&" />"&VbCrLf
End Sub
Public Sub Pop(Elem)
StackIndex=StackIndex-1
Indent()
If Elem<>Stack(StackIndex)Then
Response.Write "XML Error: Tag Mismatch when trying to close "&Elem
Else
Xml=Xml&"</"&Elem&">"&VbCrLf
End If
End Sub
Function Attribute(Key,Value)
Attribute=Key&"="""&Value&""""
End Function
Public Function GetXml()
GetXml=Xml
End Function
Private Sub Class_Terminate()
Erase Stack
End Sub
End Class
Function GetElementText(Node,Tagname)
Dim NodeList
Set NodeList=Node.getElementsByTagname(Tagname)
If NodeList.Length>0 Then
GetElementText=NodeList(0).text
Else
GetElementText=""
End If
Set NodeList=Nothing
End Function
Function GetRootNode(XmlData)
Dim XmlDOMDoc
Set XmlDOMDoc=Server.CreateObject("Msxml2.DOMDocument.3.0")
XmlDOMDoc.loadXml XmlData
Set GetRootNode=XmlDOMDoc.documentElement
Set XmlDOMDoc=Nothing
End Function
Class FWX_Azure_Redis
Public base
Public auth
Public path
Private Sub Class_Initialize()
Me.base="localhost:666"
End Sub
Public Function Key(ByVal p)
Key="https://"&Me.base&"/key?key="&Enclose(Me.path,"",".")&p
End Function
Public Function Method(ByVal m)
Method="https://"&Me.base&"/"&m
End Function
Public Function HTTP(ByVal v,ByVal u)
Dim redis
Set redis=ASyncHTTPVerb(v,u)
redis.SetRequestHeader "auth",Me.auth
Set HTTP=redis
End Function
Public Function HTTPDel(ByVal p)
HTTPDel=SendAsyncRequest(Me.HTTP("DELETE",p))
End Function
Public Function HTTPGet(ByVal p)
HTTPGet=SendAsyncRequest(Me.HTTP("GET",p))
End Function
Public Function HTTPPut(ByVal p,ByVal v)
HTTPPut=SendAsyncRequestWithData(Me.HTTP("POST",p),v)
End Function
Public Default Property Get CFG(ByVal p)
CFG=Me.Read(p)
End Property
Public Property Let CFG(ByVal p,ByVal v)
Me.Write p,v
End Property
Public Function Delete(ByVal p)
Delete=Me.HTTPDel(Me.Key(p))
End Function
Public Function Read(ByVal p)
Read=Me.HTTPGet(Me.Key(p))
End Function
Public Function Write(ByVal p,ByVal v)
If v<>"" Then
Me.HTTPPut Me.Key(p),v
Else
Me.Delete p
End If
End Function
Public Function Keys()
Keys=Me.HTTPGet(Me.Method("keys"))
End Function
Public Function Purge()
Purge=Me.HTTPDel(Me.Method("all"))
End Function
Public Property Get JSON(ByVal p)
Set JSON=TXT2Json(Me.CFG(p))
End Property
Public Property Let JSON(ByVal p,ByVal v)
Me.CFG(p)=v
End Property
End Class
Const FWX__FE_Reserved_Paths="@|x|fwx|admin|quote|cart|basket|client|members|my|your|do|login|checkout|start|pay"
Const FWX__BO_Reserved_SubDomains="((fwx|oms|cms|admin|office|staff|back|bo|driver|drivers|dp|stock|warehouse|wh)\d*)"
Dim ReqType:ReqType=Request.ServerVariables("CONTENT_TYPE")
Dim ReqLength:ReqLength=Request.ServerVariables("CONTENT_LENGTH")
Dim SITE_PATH:SITE_PATH=Server.MapPath("/")
Dim ROOT:ROOT=SITE_PATH&"\"
Dim BELOWROOT:BELOWROOT=Left(SITE_PATH,InStrRev(SITE_PATH,"\"))
Dim REFERRER:REFERRER=Request.ServerVariables("HTTP_REFERER")
Dim REFERER:REFERER=REFERRER
Dim SERVER_NAME:SERVER_NAME=Request.ServerVariables("SERVER_NAME")
Dim DOMAIN_NAME:DOMAIN_NAME=Request.ServerVariables("SERVER_NAME")
Dim SERVER_IP:SERVER_IP=Request.ServerVariables("LOCAL_ADDR")
If SERVER_IP="::1" Then SERVER_IP="127.0.0.1"
If SERVER_IP="127.0.0.1" Then SERVER_IP="79.69.67.48"
Dim CLIENT_IP:CLIENT_IP=Request.ServerVariables("HTTP_CF-Connecting-IP")
If CLIENT_IP="" Then CLIENT_IP=Request.ServerVariables("REMOTE_ADDR")
If CLIENT_IP="::1" Then CLIENT_IP="127.0.0.1"
If CLIENT_IP="127.0.0.1" Then CLIENT_IP="79.69.67.48"
Dim VISITOR_IP:VISITOR_IP=CLIENT_IP
Dim QUERY_WITH:QUERY_WITH=Request.ServerVariables("HTTP_X-Requested-With")
Dim QUERY_METHOD:QUERY_METHOD=Request.ServerVariables("HTTP_X-Method")
If QUERY_WITH="" Then QUERY_WITH=QUERY_METHOD
If QUERY_METHOD="" Then QUERY_METHOD=QUERY_METHOD
Dim QUERY_HEADERS:QUERY_HEADERS=Request.ServerVariables("ALL_HTTP")
Dim QUERY_STR:QUERY_STR=Request.ServerVariables("QUERY_STRING")
If Left(QUERY_STR,1)="?" Then QUERY_STR=Mid(QUERY_STR,2)
Dim QUERY_STRING:QUERY_STRING=QUERY_STR
Dim SCRIPT_NAME:SCRIPT_NAME=Request.ServerVariables("SCRIPT_NAME")
Dim SCRIPT:SCRIPT=SCRIPT_NAME
Dim REQUEST_METHOD:REQUEST_METHOD=Request.ServerVariables("REQUEST_METHOD")
Dim HTTP_METHOD:HTTP_METHOD=REQUEST_METHOD
Dim REQUEST_ORIGIN:REQUEST_ORIGIN=Request.ServerVariables("HTTP_ORIGIN")
If REQUEST_ORIGIN="null" Then REQUEST_ORIGIN=""
If REQUEST_ORIGIN="" And REFERRER<>"" Then REQUEST_ORIGIN=GetDomain(REFERRER)
Dim Is404:Is404=CBool(Left(QUERY_STR,4)="404;" Or Request.ServerVariables("HTTP_X_ORIGINAL_URL")<>"")
Dim IsLocal:IsLocal=CBool(SERVER_NAME="localhost" Or Left(SERVER_NAME,6)="local." Or Right(SERVER_NAME,6)=".local" Or InStr(1,"l1.l2.l3.l4.l5.l6.l7.l8.l9.l0.",Left(SERVER_NAME,3))>0 Or Left(SERVER_IP,4)="127." Or Left(SERVER_IP,4)="192.")' Or Left(CLIENT_IP,3)="192")
Dim IsLive:IsLive=CBool(Not IsLocal)
Dim IsHTTPReq:IsHTTPReq=CBool(QUERY_WITH="HTTPRequest")
Dim IsApp:IsApp=CBool(QUERY_WITH="Angular" Or QUERY_WITH="App" Or Request.ServerVariables("HTTP_X-Method")="App" Or Request.ServerVariables("HTTP_X-App")="Fyneworks")
Dim IsAjax:IsAjax=CBool(IsHTTPReq Or(QUERY_WITH="XMLHttpRequest")Or(QUERY_WITH="Ajax")Or(QUERY_WITH="Angular" Or QUERY_WITH="App" Or Request.ServerVariables("HTTP_X-Method")="Ajax"))
Dim IsSOAP:IsSOAP=CBool(InStr(1,ReqType,"multipart")>0)
Dim IsPost:IsPost=CBool(InStr(1,ReqType,"form")>0 Or ReqLength>0 Or UCase(REQUEST_METHOD)="POST")
Dim IsHTTPS:IsHTTPS=CBool(Request.ServerVariables("HTTPS")="on")
If Request.ServerVariables("HTTP_PROTOCOL")="https" Then IsHTTPS=True
If Request.ServerVariables("HTTP_X-FORWARDED-PROTO")="https" Then IsHTTPS=True
If Request.ServerVariables("HTTP_CF-VISITOR")="{""scheme"":""https""}" Then IsHTTPS=True
Dim PROTOCOL:PROTOCOL="http":If IsHTTPS Then PROTOCOL="https"
Function MyProtocol()
If PROTOCOL<>"" Then
MyProtocol=PROTOCOL
Else
If IsHTTPS Then
MyProtocol="http"
Else
MyProtocol="https"
End If
End If
End Function
Dim THIS_SERVER:THIS_SERVER=MyProtocol&"://"&SERVER_NAME
Dim FIXD_SERVER:FIXD_SERVER=MyProtocol&"://fixed-"&SERVER_NAME
Dim THIS_DOMAIN:THIS_DOMAIN=SERVER_NAME
Dim FIXD_DOMAIN:FIXD_DOMAIN="fixed-"&SERVER_NAME
Dim CACHE__timezone:CACHE__timezone=""
Dim CACHE__timezone_location:CACHE__timezone_location=""
Function MyTimezone()
If CACHE__timezone<>"" Then
MyTimezone=Trim(CACHE__timezone)
Exit Function
End If
If CACHE__timezone="" Then
If CACHE__timezone_location="" Then CACHE__timezone_location=MyBaseLocation
If CACHE__timezone_location<>"" Then
DBG "MyTimezone (by location) '"&MyBaseLocation&"'..."
CACHE__timezone=GoogleTimezoneIn(CACHE__timezone_location)
DBG "MyTimezone (by location) '"&MyBaseLocation&"' = '"&CACHE__timezone&"'"
End If
If CACHE__timezone="" And IsObject(App)Then CACHE__timezone=App.CFG("timezone")
If CACHE__timezone="" And IsObject(Application)Then CACHE__timezone=Application.Contents("timezone")
If CACHE__timezone="" And IsObject(j)Then CACHE__timezone=j("timezone")
If CACHE__timezone="" Then CACHE__timezone=My_FWX__timezone
If CACHE__timezone="" Then CACHE__timezone=FWX__timezone
If CACHE__timezone="" Then CACHE__timezone=FWX_timezone
DBG "MyTimezone (result) '"&CACHE__timezone&"'"
If CACHE__timezone<>"" Then
If IsObject(j)Then
j("TIMEZONE")=CACHE__timezone
End If
Else
CACHE__timezone=" "
End If
End If
MyTimezone=Trim(CACHE__timezone)
End Function
Dim CACHE__BaseLocation:CACHE__BaseLocation=""
Function MyBaseLocation()
If CACHE__BaseLocation<>"" Then
MyBaseLocation=Trim(CACHE__BaseLocation)
Exit Function
End If
If CACHE__BaseLocation="" Then
If CACHE__BaseLocation="" And IsObject(App)Then CACHE__BaseLocation=App.CFG("base.location")
If CACHE__BaseLocation="" And IsObject(Application)Then CACHE__BaseLocation=Application.Contents("base.location")
If CACHE__BaseLocation="" And IsObject(j)Then CACHE__BaseLocation=j("base.location")
If CACHE__BaseLocation="" Then CACHE__BaseLocation=My_FWX__server_location
If CACHE__BaseLocation="" Then CACHE__BaseLocation=FWX__server_location
If CACHE__BaseLocation="" Then CACHE__BaseLocation=FWX_server_location
If CACHE__BaseLocation="" And IsObject(shop)Then CACHE__BaseLocation=shop.CFG("city")
DBG "MyBaseLocation (result) '"&CACHE__BaseLocation&"'"
If CACHE__BaseLocation<>"" Then
If IsObject(j)Then
j("BaseLocation")=CACHE__BaseLocation
End If
Else
CACHE__BaseLocation=" "
End If
End If
MyBaseLocation=Trim(CACHE__BaseLocation)
End Function
Dim QUERY_URL:QUERY_URL=""
Dim REWRITE_URL,REWRITE_QRY
REWRITE_URL=Request.ServerVariables("HTTP_X_REWRITE_URL")
If REWRITE_URL<>"" Then
If InStr(1,REWRITE_URL,"?")>0 Then
REWRITE_QRY=Mid(REWRITE_URL,InStr(1,REWRITE_URL,"?")+1)
REWRITE_URL=Left(REWRITE_URL,InStr(1,REWRITE_URL,"?")-1)
End If
QUERY_URL=REWRITE_URL
QUERY_STR=REWRITE_QRY
End If
Dim ORIGINAL_URL,ORIGINAL_QRY
ORIGINAL_URL=Request.ServerVariables("HTTP_X_ORIGINAL_URL")
If ORIGINAL_URL<>"" Then
If InStr(1,ORIGINAL_URL,"?")>0 Then
ORIGINAL_QRY=Mid(ORIGINAL_URL,InStr(1,ORIGINAL_URL,"?")+1)
ORIGINAL_URL=Left(ORIGINAL_URL,InStr(1,ORIGINAL_URL,"?")-1)
End If
QUERY_URL=ORIGINAL_URL
QUERY_STR=ORIGINAL_QRY
End If
If QUERY_URL="" Then
If Is404 Then
QUERY_URL=Mid(QUERY_STR,InStr(1,QUERY_STR,SERVER_NAME)+Len(SERVER_NAME))
If InStr(1,QUERY_URL,"?")>0 Then
QUERY_URL=Mid(QUERY_URL,1,InStr(1,QUERY_URL,"?")-1)
End If
QUERY_URL=Replace(QUERY_URL,":80","")
QUERY_URL=Replace(QUERY_URL,":443","")
If InStr(1,QUERY_URL,"?")>0 Then
QUERY_STR=Mid(QUERY_URL,InStr(1,QUERY_URL,"?")+1)
Else
QUERY_STR=""
End If
End If
End If
If QUERY_URL="" Then QUERY_URL=SCRIPT_NAME
Dim QUERY_PATH:QUERY_PATH=Left(QUERY_URL,InStrRev(QUERY_URL,"/"))
Dim QUERY_URI:QUERY_URI=QUERY_URL
If QUERY_STR<>"" Then QUERY_URI=QUERY_URI&"?"
QUERY_URI=QUERY_URI&QUERY_STR
Dim j:Set j=Dictionary()
Dim TIMESTAMP:TIMESTAMP=RemoveSymbols(Timer())
j("DATESTAMP")=Now()
j("STARTFLAG")=Timer()
j("TIMESTAMP")=TIMESTAMP
j("NOW")=FmtDate(NewLocalDate("now"),"%Y-%M-%D %H:%N:%S")
j("DATE")=FmtDate(j("NOW"),"%Y-%M-%D")
j("TODAY")=j("DATE")
j("URL")=QUERY_URL
j("QRY")=QUERY_STR
If QUERY_STR<>"" Then j("qid")=UCase(DoEncrypt(MD5(QUERY_STR),"FYNEWORKS"))
If QUERY_URL<>"" Then j("uid")=UCase(DoEncrypt(MD5(QUERY_URL),"WORKSFINE"))
If QUERY_URI<>"" Then j("fid")=UCase(DoEncrypt(MD5(QUERY_URI),"EVERYDAYS"))
j("uuid")="T"&j("TIMESTAMP")&"-"&j("uid")
If j("qid")="EOG9NJGEI" Then
j("qid")=""
QUERY_STR=""
End If
j("ip")=VISITOR_IP
j("human")=IIf(IsHuman,True,False)
j("cur")=TOT(Application.Contents("cur"),"GBP")
Const FWX__EMPTY="<!--EMPTY-->"
Dim IsBack:IsBack=CBool(rxTest(QUERY_URL,"^/(x|admin|fwx)\b/?"))
Dim IsFront:IsFront=CBool(Not IsBack)
Dim USERAGENT:USERAGENT=Request.ServerVariables("HTTP_USER_AGENT")
Dim IsIE:IsIE=CBool(InStr(1,USERAGENT,"MSIE")>0)'rxTest(USERAGENT,"MSIE")
Dim IsFF:IsFF=CBool(InStr(1,USERAGENT,"Gecko")>0)'rxTest(USERAGENT,"Gecko")
Dim IsWK:IsWK=CBool(InStr(1,USERAGENT,"Webkit")>0)'rxTest(USERAGENT,"Webkit")
Dim IsOpera:IsOpera=CBool(InStr(1,USERAGENT,"Opera")>0)'rxTest(USERAGENT,"Opera")
Dim IsChrome:IsChrome=CBool(InStr(1,USERAGENT,"Chrome")>0)'rxTest(USERAGENT,"Chrome")
Dim IsSafari:IsSafari=CBool(InStr(1,USERAGENT,"Webkit")>0)'rxTest(USERAGENT,"Webkit")
Dim BrowserCode:BrowserCode=IIf(IsIE,"IE","FF")
Dim IEVersion:IEVersion=0
If Not IsIE Then
ElseIf InStr(1,USERAGENT,"MSIE 8")>0 Then:IEVersion=8
ElseIf InStr(1,USERAGENT,"MSIE 7")>0 Then:IEVersion=7
ElseIf InStr(1,USERAGENT,"MSIE 6")>0 Then:IEVersion=6
Else:IEVersion=5
End If
Dim IsGoogle:IsGoogle=CBool(InStr(1,USERAGENT,"Google")>0)
Dim IsYahoo:IsYahoo=CBool(InStr(1,USERAGENT,"Slurp")>0)
Dim IsTeoma:IsTeoma=CBool(InStr(1,USERAGENT,"MSNBOT")>0)
Dim IsAlexa:IsAlexa=CBool(InStr(1,USERAGENT,"ia_archiver")>0)
Dim IsAltavista:IsAltavista=CBool(InStr(1,USERAGENT,"Scooter")>0 Or InStr(1,USERAGENT,"Mercator")>0)
Dim IsAllTheWeb:IsAllTheWeb=CBool(InStr(1,USERAGENT,"FAST")>0)
Dim IsLookSmart:IsLookSmart=CBool(InStr(1,USERAGENT,"MantraAgent")>0)
Dim IsLycos:IsLycos=CBool(InStr(1,USERAGENT,"Lycos")>0)
Dim IsWISEnut:IsWISEnut=CBool(InStr(1,USERAGENT,"ZyBorg")>0)
Const GoogleAPI="JBcqdp9QFHItpcifgj8WvI1KEkdK9CKe"'for "diego@fyneworks.com"
Const qryNavVars="q&@size&@page"
Const qryClearNavVars="&~q&~@size&~@page"
Const qryClearAssetVars="&~download&~page&~image&~link&~mode&"
Const fwxDefSite=-9
Const fwxDemoSite=-3
Const fwxHelpSite=-2
Const fwxSysSite=-1
Const TAG_OPEN="["
Const TAG_CLOSE="]"
Const fwxMasterAccess="[ogeid]"
Const fwxStaffAccess="[staff][admin][manager][ceo][editor][clerk]"
j("country_restriction")=""
j("country_billing")=TOT(j("country_billing"),"UK")
j("country_billing_restriction")=""
Function FWX_Country_Billing()
FWX_Country_Billing=TOT(j("country_billing"),j("country"))
End Function
Function FWX_Country_Billing_Restriction_Country()
Dim o
o=TOT(o,j("country_billing_restriction"))
o=TOT(o,j("country_restriction_billing"))
o=TOT(o,j("country_restriction"))
FWX_Country_Billing_Restriction_Country=o
End Function
Function FWX_Country_Billing_Restriction()
Dim o
o=FWX_Country_Billing_Restriction_Country
If o<>"" Then
FWX_Country_Billing_Restriction="dbo.InList(t.code,'"&o&"')=1"
End If
End Function
j("country_shipping")=TOT(j("country_shipping"),"UK")
j("country_shipping_restriction")="UK"
Function FWX_Country_Shipping()
FWX_Country_Shipping=TOT(j("country_shipping"),j("country_billing"))
End Function
Function FWX_Country_Shipping_Restriction_Country()
Dim o
o=TOT(o,j("country_shipping_restriction"))
o=TOT(o,j("country_restriction_shipping"))
o=TOT(o,j("country_restriction"))
FWX_Country_Shipping_Restriction_Country=o
End Function
Function FWX_Country_Shipping_Restriction()
Dim o
o=FWX_Country_Shipping_Restriction_Country
If o<>"" Then
FWX_Country_Shipping_Restriction="dbo.InList(t.code,'"&o&"')=1"
End If
End Function
Dim IsMobileDevice,IsMobile,IsDesktop
IsMobileDevice=IsMobileCheck()
MobileDeviceOverride="n"'MyCookies("Wants.Mobile")
If MobileDeviceOverride<>"" Then
IsMobile=CBool(MobileDeviceOverride="y")
IsDesktop=(Not IsMobile)
Else
IsMobile=IsMobileDevice
IsDesktop=(Not IsMobile)
End If
Dim IsCDN,IsTemp,IsTest,IsBOSub
Dim LAST_DOMAIN_NAME
LAST_DOMAIN_NAME=MyDomain
EXE "FWX_DOMAIN__PARSE"
EXE "FWX_CONSTANTS"
If DOMAIN_NAME<>LAST_DOMAIN_NAME Then
EXE "FWX_DOMAIN__PARSE"
LAST_DOMAIN_NAME=DOMAIN_NAME
End If
EXE "FWX_DOMAIN__SET"
If DOMAIN_NAME<>LAST_DOMAIN_NAME Then
EXE "FWX_DOMAIN__PARSE"
LAST_DOMAIN_NAME=DOMAIN_NAME
End If
Dim MyDomainTLD,MyDomainTLDLess,IsSubDomain,SUBDOMAIN
MyDomainTLD=GetTLD(MyDomain)
MyDomainTLDLess=GetTLDLess(MyDomain)
IsSubDomain=CBool(InStr(1,MyDomainTLDLess,".")>0)
Dim IsTestBot:IsTestBot=rxTest(USERAGENT,"Testomato|UptimeBot|Uptimerobot|Screaming Frog")
Dim IsHuman:IsHuman=CBool(IsIE Or IsFF Or IsWK Or IsNS Or IsChrome Or IsOpera Or IsMobile Or IsMobileDevice)
Dim IsRobot:IsRobot=CBool(Not IsHuman)
If IsTestBot Or rxTest(USERAGENT,"baidu")Then
IsRobot=True
IsHuman=False
End If
Dim IsBot:IsBot=IsRobot
j("MySite")=MySite
j("server")=SERVER_NAME
j("domain")=IIf(IsLocal,SERVER_NAME,MyDomain)
j("www.domain")=IIf(IsLocal,SERVER_NAME,MyWWWDomain)
j("sitePath")=TMPL_FOLDER'"/$"&Replace(MyDomain,".","-")&"/"
j("dataPath")=TMPL_FOLDER'j("sitePath")&""
j("TMPL_FOLDER")=TMPL_FOLDER
j("SITE_FOLDER")=SITE_FOLDER
Function FWX_DOMAIN__PARSE()
DOMAIN_NAME=rxReplace(DOMAIN_NAME,"^www\.","")
If Not IsLocal Then
IsLocal=rxTest(DOMAIN_NAME,"^(local|l\d)\.")Or(Not InStr(1,SERVER_NAME,".")>0)
End If
If Not IsLocal Then
For Each tester In Split("lvh.me,vcap.me,fuf.me,ulh.us",",")
If Not IsLocal Then
tester="\."&Replace(tester,".","\.")&"$"
If rxTest(DOMAIN_NAME,tester)Then
IsLocal=True
DOMAIN_NAME=rxReplace(DOMAIN_NAME,tester,"")
End If
End If
Next
End If
DOMAIN_NAME=rxReplace(DOMAIN_NAME,"^(local|l\d)\.","")
IsLive=(Not IsLocal)
IsTemp=CBool(IsTemp Or rxTest(DOMAIN_NAME,"\b(temp|test|dev|d\d|svn|predns|ourhelmcp|ourwindowsnetwork|azurewebsites)\b"))
IsTest=CBool(IsTest Or CBool(IsLocal Or IsTemp Or IsTrue(Application.Contents("DevMode"))))
If IsTest Or IsTemp Then
DOMAIN_NAME=rxReplace(DOMAIN_NAME,"(\.(temp|predns))?\.?(ourhelmcp|ourwindowsnetwork)\.com$","")
DOMAIN_NAME=rxReplace(DOMAIN_NAME,"\b(temp|test|dev|d\d|d\d|svn)\.?","")
End If
If Not IsCDN Then IsCDN=rxTest(DOMAIN_NAME,"^w\d*\.")
If IsCDN Then DOMAIN_NAME=rxReplace(DOMAIN_NAME,"^w\d*\.","")
If rxTest(DOMAIN_NAME,"^((l\d|local|test|temp|dev)\.)?("&FWX__BO_Reserved_SubDomains&")\.")Then
IsBOSub=True
IsSubDomain=True
SUBDOMAIN=rxSelect(rxSelect(DOMAIN_NAME,"^((l\d|local|test|temp|dev)\.)?("&FWX__BO_Reserved_SubDomains&")\b"),"\w+$")
DOMAIN_NAME=rxReplace(DOMAIN_NAME,"^((l\d|local|test|temp|dev)\.)?("&FWX__BO_Reserved_SubDomains&")\.","")
End If
If rxTest(DOMAIN_NAME,"^m(obi)?\.")Then
DOMAIN_NAME=rxReplace(DOMAIN_NAME,"^m(ob|obi)?\.","")
If True Then
PermanentRedirect "http://www."&DOMAIN_NAME&rxReplace(QUERY_URL,"/$","")
End If
If MyCookies("Wants.Mobile")="" Then
IsMobile=True
IsDesktop=False
End If
End If
If rxTest(DOMAIN_NAME,"^d(esktop)?\.")Then
DOMAIN_NAME=rxReplace(DOMAIN_NAME,"^m(ob|obi)?\.","")
If MyCookies("Wants.Mobile")="" Then
IsMobile=False
IsDesktop=True
End If
End If
MyDomainTLDLess=rxReplace(MyDomain,"(\.\w{2,3})?(\.\w{2,3})$","")
IsSubDomain=CBool(InStr(1,MyDomainTLDLess,".")>0)
End Function
Function IsIPad()
IsiPad=rxTest(USERAGENT,"iPad")
End Function
Function IsMobileCheck()
IsMobileCheck=False
If IsIPad Then
ElseIf Request.ServerVariables("HTTP_X_WAP_PROFILE")<>"" Then
IsMobileCheck=True
ElseIf rxTest(Request.ServerVariables("HTTP_ACCEPT"),"wap\.|\.wap")Then
IsMobileCheck=True
ElseIf rxTest(USERAGENT,Join(Array("OfficeLiveConnector","MSIE\ 8\.0","OptimizedIE8","MSN\ Optimized","Creative\ AutoUpdate","Swapper"),"|"))Then
ElseIf rxTest(USERAGENT,Join(Array("Nexus 4","Nexus 5","Mobile Safari"),"|"))Then
IsMobileCheck=True
ElseIf rxTest(USERAGENT,Join(Array("midp","j2me","avantg","docomo","novarra","palmos","palmsource","240x320","opwv","chtml","pda","windows\ ce","mmp\/","blackberry","mib\/","symbian","wireless","nokia","hand","mobi","phone","cdm","up\.b","audio","SIE\-","SEC\-","samsung","HTC","mot\-","mitsu","sagem","sony","alcatel","lg","erics","vx","NEC","philips","mmm","xx","panasonic","sharp","wap","sch","rover","pocket","benq","java","pt","pg","vox","amoi","bird","compal","kg","voda","sany","kdd","dbt","sendo","sgh","gradi","jb","\d\d\di","moto","webos"),"|"))Then
IsMobileCheck=True
End If
End Function
Const rxTime="^(\d\d):(\d\d)(:(\d\d))?"
Const rxTimeOnly="^(\d\d):(\d\d)"
Const fwxDate="%D-%b-%Y"
Const fwxLongDate="%d-%b-%y @ %H:%N"
Const fwxShortDate="%d-%b-%y"
Const fwxTime="%H:%N"
Const fwxDateOnly="%D-%b-%Y"
Const fwxTimeOnly="%H:%N"
Const fwxDateTime="%D-%b-%Y %H:%N:%S"
Const fwxDateOrdinal="%Y%M%D%H%N%S"
fwxDateNumbers=fwxDateOrdinal
Const fwxDateWin="%D-%b-%Y %H:%N:%S"
Const fwxDateXml="%a, %D %b %Y %H:%N:%S GMT"
Const fwxDateGoogle="%Y-%M-%DT%H:%N:%S+00:00"
Const fwxDateRSS="%a, %D %b %Y %H:%N:%S GMT"
Const fwxDateRDF="%Y-%M-%DT%H:%N:%S+00:00"
Const fwxDateAtom="%Y-%M-%DT%H:%N:%S+00:00"
Const fwxDateROR="%Y-%M-%DT%H:%N:%S+00:00"
Const fwxDateSQL="%Y-%M-%DT%H%N%SZ+00:00"
Const fwxDateFSO="%Y-%M-%DT%H%N%S"
Const fwxDateFile="%Y-%M-%DT%H%N%S"
Const fwxDateYMD="%Y%M%D"
Const fwxDateRFC822="%a, %D %b %Y %H:%N:%S GMT"
Const fwxDateRFC3339="%Y-%M-%DT%H:%N:%S+00:00"
fwxDateRFC=fwxDateRFC3339
Const fwxDateISO8601="%Y-%M-%DT%H:%N:%SZ+00:00"
fwxDateISO=fwxDateISO8601
Const fwxDateICS="%Y%M%DT%H%N%SZ"
Const fwxDateSearch=",t.date:year:c-equal:s-year:alt-yyyy,t.date:yy:c-equal:s-yy,t.date:month:c-equal:s-month,t.date:day:c-equal:s-day,t.date:week:c-equal:s-week,t.date:weekday:c-equal:s-weekday,t.date:date:c-equal:s-date,t.date:date:c-moreequal:s-after:alt-start,t.date:date:c-lessequal:s-before:alt-end,"
Const ALPHABET="abcdefghijklmnopqrstuvxyz0123456789"
Const rxASCII="[\x09\x0A\x0D\x20-\x7E\xA0-\xFF]"
Const rxNonASCII="[^\x09\x0A\x0D\x20-\x7E\xA0-\xFF]"
Const fwxMonthNames="January|February|March|April|May|June|July|August|September|October|November|December"
Const fwxMonthAbbrev="Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
Dim rxMonthName:rxMonthName="(("&fwxMonthNames&")|("&fwxMonthAbbrev&"))"
Const fwxDayNames="Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday"
Const fwxDayAbbrev="Sun|Mon|Tue|Wed|Thu|Fri|Sat"
Dim rxDayName:rxDayName="(("&fwxDayNames&")|("&fwxDayAbbrev&"))"
Const fwxCountry="UK"
Const fwxCountry_Europe="aa|an|at|be|bu|cs|dk|fi|fr|gb|gw|gi|gr|hu|ic|ie|it|lh|lu|mm|mc|ne|no|pl|po|rm|sm|sp|sw|sz|uk|vc|wb|yu|gb"
Const FWX_Cur_ALL="&#x4c;&#x65;&#x6b;"
Const FWX_Cur_AFN="&#x60b;"
Const FWX_Cur_ARS="&#x24;"
Const FWX_Cur_AWG="&#x192;"
Const FWX_Cur_AUD="&#x24;"
Const FWX_Cur_AZN="&#x43c;&#x430;&#x43d;"
Const FWX_Cur_BSD="&#x24;"
Const FWX_Cur_BBD="&#x24;"
Const FWX_Cur_BYR="&#x70;&#x2e;"
Const FWX_Cur_BZD="&#x42;&#x5a;&#x24;"
Const FWX_Cur_BMD="&#x24;"
Const FWX_Cur_BOB="&#x24;&#x62;"
Const FWX_Cur_BAM="&#x4b;&#x4d;"
Const FWX_Cur_BWP="&#x50;"
Const FWX_Cur_BGN="&#x43b;&#x432;"
Const FWX_Cur_BRL="&#x52;&#x24;"
Const FWX_Cur_BND="&#x24;"
Const FWX_Cur_KHR="&#x17db;"
Const FWX_Cur_CAD="&#x24;"
Const FWX_Cur_KYD="&#x24;"
Const FWX_Cur_CLP="&#x24;"
Const FWX_Cur_CNY="&#xa5;"
Const FWX_Cur_COP="&#x24;"
Const FWX_Cur_CRC="&#x20a1;"
Const FWX_Cur_HRK="&#x6b;&#x6e;"
Const FWX_Cur_CUP="&#x20b1;"
Const FWX_Cur_CZK="&#x4b;&#x10d;"
Const FWX_Cur_DKK="&#x6b;&#x72;"
Const FWX_Cur_DOP="&#x52;&#x44;&#x24;"
Const FWX_Cur_XCD="&#x24;"
Const FWX_Cur_EGP="&#xa3;"
Const FWX_Cur_SVC="&#x24;"
Const FWX_Cur_EEK="&#x6b;&#x72;"
Const FWX_Cur_EUR="&#x20ac;"
Const FWX_Cur_FKP="&#xa3;"
Const FWX_Cur_FJD="&#x24;"
Const FWX_Cur_GHC="&#xa2;"
Const FWX_Cur_GIP="&#xa3;"
Const FWX_Cur_GTQ="&#x51;"
Const FWX_Cur_GGP="&#xa3;"
Const FWX_Cur_GYD="&#x24;"
Const FWX_Cur_HNL="&#x4c;"
Const FWX_Cur_HKD="&#x24;"
Const FWX_Cur_HUF="&#x46;&#x74;"
Const FWX_Cur_ISK="&#x6b;&#x72;"
Const FWX_Cur_INR="&#x20b9;"
Const FWX_Cur_IDR="&#x52;&#x70;"
Const FWX_Cur_IRR="&#xfdfc;"
Const FWX_Cur_IMP="&#xa3;"
Const FWX_Cur_ILS="&#x20aa;"
Const FWX_Cur_JMD="&#x4a;&#x24;"
Const FWX_Cur_JPY="&#xa5;"
Const FWX_Cur_JEP="&#xa3;"
Const FWX_Cur_KZT="&#x43b;&#x432;"
Const FWX_Cur_KGS="&#x43b;&#x432;"
Const FWX_Cur_LAK="&#x20ad;"
Const FWX_Cur_LVL="&#x4c;&#x73;"
Const FWX_Cur_LBP="&#xa3;"
Const FWX_Cur_LRD="&#x24;"
Const FWX_Cur_LTL="&#x4c;&#x74;"
Const FWX_Cur_MKD="&#x434;&#x435;&#x43d;"
Const FWX_Cur_MYR="&#x52;&#x4d;"
Const FWX_Cur_MUR="&#x20a8;"
Const FWX_Cur_MXN="&#x24;"
Const FWX_Cur_MNT="&#x20ae;"
Const FWX_Cur_MZN="&#x4d;&#x54;"
Const FWX_Cur_NAD="&#x24;"
Const FWX_Cur_NPR="&#x20a8;"
Const FWX_Cur_ANG="&#x192;"
Const FWX_Cur_NZD="&#x24;"
Const FWX_Cur_NIO="&#x43;&#x24;"
Const FWX_Cur_NGN="&#x20a6;"
Const FWX_Cur_KPW="&#x20a9;"
Const FWX_Cur_NOK="&#x6b;&#x72;"
Const FWX_Cur_OMR="&#xfdfc;"
Const FWX_Cur_PKR="&#x20a8;"
Const FWX_Cur_PAB="&#x42;&#x2f;&#x2e;"
Const FWX_Cur_PYG="&#x47;&#x73;"
Const FWX_Cur_PEN="&#x53;&#x2f;&#x2e;"
Const FWX_Cur_PHP="&#x20b1;"
Const FWX_Cur_PLN="&#x7a;&#x142;"
Const FWX_Cur_QAR="QR" ' رق 'QR '&#1583;&#1573;" '"&#xfdfc;" '"ر.ق" '"QR" 'ر.ق" '"ر.ق" "QAR" '"&#xfdfc;"
Const FWX_Cur_QAR_Decimals=0
Const FWX_Cur_RON="&#x6c;&#x65;&#x69;"
Const FWX_Cur_RUB="&#x440;&#x443;&#x431;"
Const FWX_Cur_SHP="&#xa3;"
Const FWX_Cur_SAR="&#xfdfc;"
Const FWX_Cur_RSD="&#x414;&#x438;&#x43d;&#x2e;"
Const FWX_Cur_SCR="&#x20a8;"
Const FWX_Cur_SGD="&#x24;"
Const FWX_Cur_SBD="&#x24;"
Const FWX_Cur_SOS="&#x53;"
Const FWX_Cur_ZAR="&#x52;"
Const FWX_Cur_KRW="&#x20a9;"
Const FWX_Cur_LKR="&#x20a8;"
Const FWX_Cur_SEK="&#x6b;&#x72;"
Const FWX_Cur_CHF="&#x43;&#x48;&#x46;"
Const FWX_Cur_SRD="&#x24;"
Const FWX_Cur_SYP="&#xa3;"
Const FWX_Cur_TWD="&#x4e;&#x54;&#x24;"
Const FWX_Cur_THB="&#xe3f;"
Const FWX_Cur_TTD="&#x54;&#x54;&#x24;"
Const FWX_Cur_TRL="&#x20a4;"
Const FWX_Cur_TVD="&#x24;"
Const FWX_Cur_UAH="&#x20b4;"
Const FWX_Cur_AED="&#1583;&#1573;"
Const FWX_Cur_AED_Decimals=0
Const FWX_Cur_GBP="&#xa3;"
Const FWX_Cur_GBP_Decimals=2
Const FWX_Cur_USD="&#x24;"
Const FWX_Cur_USD_Decimals=2
Const FWX_Cur_UYU="&#x24;&#x55;"
Const FWX_Cur_UZS="&#x43b;&#x432;"
Const FWX_Cur_VEF="&#x42;&#x73;"
Const FWX_Cur_VND="&#x20ab;"
Const FWX_Cur_YER="&#xfdfc;"
Const FWX_Cur_ZWD="&#x5a;&#x24;"
Function Country00(ByVal s)
Dim d:Set d=CountryData(s)
Country00=TOT(d("00"),s)
Set d=Nothing
End Function
Function CountryXX(ByVal s)
Dim d:Set d=CountryData(s)
CountryXX=TOT(d("xx"),s)
Set d=Nothing
End Function
Function CountryXXX(ByVal s)
Dim d:Set d=CountryData(s)
CountryXXX=TOT(d("xxx"),s)
Set d=Nothing
End Function
Function CountryData(ByVal s)
Dim x,c,a,rC
Set x=Dictionary()
If GTZ(s)Then
Set rC=app.Countries.ByID(VOZ(s))
If Exists(rC)Then
s=""
UpdateCollection x,rC
End If
KillRS rC
End If
If s<>"" Then
Select Case UCase(TOE(s))
Case "93","AF","AFG":c="93:AF:AFG:Afghanistan"
Case "355","AL","ALB":c="355:AL:ALB:Albania"
Case "213","DZ","DZA":c="213:DZ:DZA:Algeria"
Case "1-684","AS","ASM":c="1-684:AS:ASM:American Samoa"
Case "376","AD","AND":c="376:AD:AND:Andorra"
Case "244","AO","AGO":c="244:AO:AGO:Angola"
Case "1-264","AI","AIA":c="1-264:AI:AIA:Anguilla"
Case "672","AQ","ATA":c="672:AQ:ATA:Antarctica"
Case "1-268","AG","ATG":c="1-268:AG:ATG:Antigua and Barbuda"
Case "54","AR","ARG":c="54:AR:ARG:Argentina"
Case "374","AM","ARM":c="374:AM:ARM:Armenia"
Case "297","AW","ABW":c="297:AW:ABW:Aruba"
Case "61","AU","AUS":c="61:AU:AUS:Australia"
Case "43","AT","AUT":c="43:AT:AUT:Austria"
Case "994","AZ","AZE":c="994:AZ:AZE:Azerbaijan"
Case "1-242","BS","BHS":c="1-242:BS:BHS:Bahamas"
Case "973","BH","BHR":c="973:BH:BHR:Bahrain"
Case "880","BD","BGD":c="880:BD:BGD:Bangladesh"
Case "1-246","BB","BRB":c="1-246:BB:BRB:Barbados"
Case "375","BY","BLR":c="375:BY:BLR:Belarus"
Case "32","BE","BEL":c="32:BE:BEL:Belgium"
Case "501","BZ","BLZ":c="501:BZ:BLZ:Belize"
Case "229","BJ","BEN":c="229:BJ:BEN:Benin"
Case "1-441","BM","BMU":c="1-441:BM:BMU:Bermuda"
Case "975","BT","BTN":c="975:BT:BTN:Bhutan"
Case "591","BO","BOL":c="591:BO:BOL:Bolivia"
Case "387","BA","BIH":c="387:BA:BIH:Bosnia and Herzegovina"
Case "267","BW","BWA":c="267:BW:BWA:Botswana"
Case "55","BR","BRA":c="55:BR:BRA:Brazil"
Case "246","IO","IOT":c="246:IO:IOT:British Indian Ocean Territory"
Case "1-284","VG","VGB":c="1-284:VG:VGB:British Virgin Islands"
Case "673","BN","BRN":c="673:BN:BRN:Brunei"
Case "359","BG","BGR":c="359:BG:BGR:Bulgaria"
Case "226","BF","BFA":c="226:BF:BFA:Burkina Faso"
Case "95","MM","MMR":c="95:MM:MMR:Myanmar"
Case "257","BI","BDI":c="257:BI:BDI:Burundi"
Case "855","KH","KHM":c="855:KH:KHM:Cambodia"
Case "237","CM","CMR":c="237:CM:CMR:Cameroon"
Case "1","CA","CAN":c="1:CA:CAN:Canada"
Case "238","CV","CPV":c="238:CV:CPV:Cape Verde"
Case "1-345","KY","CYM":c="1-345:KY:CYM:Cayman Islands"
Case "236","CF","CAF":c="236:CF:CAF:Central African Republic"
Case "235","TD","TCD":c="235:TD:TCD:Chad"
Case "56","CL","CHL":c="56:CL:CHL:Chile"
Case "86","CN","CHN":c="86:CN:CHN:China"
Case "61","CX","CXR":c="61:CX:CXR:Christmas Island"
Case "61","CC","CCK":c="61:CC:CCK:Cocos Islands"
Case "57","CO","COL":c="57:CO:COL:Colombia"
Case "269","KM","COM":c="269:KM:COM:Comoros"
Case "242","CG","COG":c="242:CG:COG:Republic of the Congo"
Case "243","CD","COD":c="243:CD:COD:Democratic Republic of the Congo"
Case "682","CK","COK":c="682:CK:COK:Cook Islands"
Case "506","CR","CRI":c="506:CR:CRI:Costa Rica"
Case "385","HR","HRV":c="385:HR:HRV:Croatia"
Case "53","CU","CUB":c="53:CU:CUB:Cuba"
Case "599","CW","CUW":c="599:CW:CUW:Curacao"
Case "357","CY","CYP":c="357:CY:CYP:Cyprus"
Case "420","CZ","CZE":c="420:CZ:CZE:Czech Republic"
Case "45","DK","DNK":c="45:DK:DNK:Denmark"
Case "253","DJ","DJI":c="253:DJ:DJI:Djibouti"
Case "1-767","DM","DMA":c="1-767:DM:DMA:Dominica"
Case "1-809","1-829","1-849","DO","DOM":c="1-849:DO:DOM:Dominican Republic"
Case "670","TL","TLS":c="670:TL:TLS:East Timor"
Case "593","EC","ECU":c="593:EC:ECU:Ecuador"
Case "20","EG","EGY":c="20:EG:EGY:Egypt"
Case "503","SV","SLV":c="503:SV:SLV:El Salvador"
Case "240","GQ","GNQ":c="240:GQ:GNQ:Equatorial Guinea"
Case "291","ER","ERI":c="291:ER:ERI:Eritrea"
Case "372","EE","EST":c="372:EE:EST:Estonia"
Case "251","ET","ETH":c="251:ET:ETH:Ethiopia"
Case "500","FK","FLK":c="500:FK:FLK:Falkland Islands"
Case "298","FO","FRO":c="298:FO:FRO:Faroe Islands"
Case "679","FJ","FJI":c="679:FJ:FJI:Fiji"
Case "358","FI","FIN":c="358:FI:FIN:Finland"
Case "33","FR","FRA":c="33:FR:FRA:France"
Case "689","PF","PYF":c="689:PF:PYF:French Polynesia"
Case "241","GA","GAB":c="241:GA:GAB:Gabon"
Case "220","GM","GMB":c="220:GM:GMB:Gambia"
Case "995","GE","GEO":c="995:GE:GEO:Georgia"
Case "49","DE","DEU":c="49:DE:DEU:Germany"
Case "233","GH","GHA":c="233:GH:GHA:Ghana"
Case "350","GI","GIB":c="350:GI:GIB:Gibraltar"
Case "30","GR","GRC":c="30:GR:GRC:Greece"
Case "299","GL","GRL":c="299:GL:GRL:Greenland"
Case "1-473","GD","GRD":c="1-473:GD:GRD:Grenada"
Case "1-671","GU","GUM":c="1-671:GU:GUM:Guam"
Case "502","GT","GTM":c="502:GT:GTM:Guatemala"
Case "44-1481","GG","GGY":c="44-1481:GG:GGY:Guernsey"
Case "224","GN","GIN":c="224:GN:GIN:Guinea"
Case "245","GW","GNB":c="245:GW:GNB:Guinea-Bissau"
Case "592","GY","GUY":c="592:GY:GUY:Guyana"
Case "509","HT","HTI":c="509:HT:HTI:Haiti"
Case "504","HN","HND":c="504:HN:HND:Honduras"
Case "852","HK","HKG":c="852:HK:HKG:Hong Kong"
Case "36","HU","HUN":c="36:HU:HUN:Hungary"
Case "354","IS","ISL":c="354:IS:ISL:Iceland"
Case "91","IN","IND":c="91:IN:IND:India"
Case "62","ID","IDN":c="62:ID:IDN:Indonesia"
Case "98","IR","IRN":c="98:IR:IRN:Iran"
Case "964","IQ","IRQ":c="964:IQ:IRQ:Iraq"
Case "353","IE","IRL":c="353:IE:IRL:Ireland"
Case "44-1624","IM","IMN":c="44-1624:IM:IMN:Isle of Man"
Case "972","IL","ISR":c="972:IL:ISR:Israel"
Case "39","IT","ITA":c="39:IT:ITA:Italy"
Case "225","CI","CIV":c="225:CI:CIV:Ivory Coast"
Case "1-876","JM","JAM":c="1-876:JM:JAM:Jamaica"
Case "81","JP","JPN":c="81:JP:JPN:Japan"
Case "44-1534","JE","JEY":c="44-1534:JE:JEY:Jersey"
Case "962","JO","JOR":c="962:JO:JOR:Jordan"
Case "7","KZ","KAZ":c="7:KZ:KAZ:Kazakhstan"
Case "254","KE","KEN":c="254:KE:KEN:Kenya"
Case "686","KI","KIR":c="686:KI:KIR:Kiribati"
Case "383","XK","XKX":c="383:XK:XKX:Kosovo"
Case "965","KW","KWT":c="965:KW:KWT:Kuwait"
Case "996","KG","KGZ":c="996:KG:KGZ:Kyrgyzstan"
Case "856","LA","LAO":c="856:LA:LAO:Laos"
Case "371","LV","LVA":c="371:LV:LVA:Latvia"
Case "961","LB","LBN":c="961:LB:LBN:Lebanon"
Case "266","LS","LSO":c="266:LS:LSO:Lesotho"
Case "231","LR","LBR":c="231:LR:LBR:Liberia"
Case "218","LY","LBY":c="218:LY:LBY:Libya"
Case "423","LI","LIE":c="423:LI:LIE:Liechtenstein"
Case "370","LT","LTU":c="370:LT:LTU:Lithuania"
Case "352","LU","LUX":c="352:LU:LUX:Luxembourg"
Case "853","MO","MAC":c="853:MO:MAC:Macao"
Case "389","MK","MKD":c="389:MK:MKD:Macedonia"
Case "261","MG","MDG":c="261:MG:MDG:Madagascar"
Case "265","MW","MWI":c="265:MW:MWI:Malawi"
Case "60","MY","MYS":c="60:MY:MYS:Malaysia"
Case "960","MV","MDV":c="960:MV:MDV:Maldives"
Case "223","ML","MLI":c="223:ML:MLI:Mali"
Case "356","MT","MLT":c="356:MT:MLT:Malta"
Case "692","MH","MHL":c="692:MH:MHL:Marshall Islands"
Case "222","MR","MRT":c="222:MR:MRT:Mauritania"
Case "230","MU","MUS":c="230:MU:MUS:Mauritius"
Case "262","YT","MYT":c="262:YT:MYT:Mayotte"
Case "52","MX","MEX":c="52:MX:MEX:Mexico"
Case "691","FM","FSM":c="691:FM:FSM:Micronesia"
Case "373","MD","MDA":c="373:MD:MDA:Moldova"
Case "377","MC","MCO":c="377:MC:MCO:Monaco"
Case "976","MN","MNG":c="976:MN:MNG:Mongolia"
Case "382","ME","MNE":c="382:ME:MNE:Montenegro"
Case "1-664","MS","MSR":c="1-664:MS:MSR:Montserrat"
Case "212","MA","MAR":c="212:MA:MAR:Morocco"
Case "258","MZ","MOZ":c="258:MZ:MOZ:Mozambique"
Case "264","NA","NAM":c="264:NA:NAM:Namibia"
Case "674","NR","NRU":c="674:NR:NRU:Nauru"
Case "977","NP","NPL":c="977:NP:NPL:Nepal"
Case "31","NL","NLD":c="31:NL:NLD:Netherlands"
Case "599","AN","ANT":c="599:AN:ANT:Netherlands Antilles"
Case "687","NC","NCL":c="687:NC:NCL:New Caledonia"
Case "64","NZ","NZL":c="64:NZ:NZL:New Zealand"
Case "505","NI","NIC":c="505:NI:NIC:Nicaragua"
Case "227","NE","NER":c="227:NE:NER:Niger"
Case "234","NG","NGA":c="234:NG:NGA:Nigeria"
Case "683","NU","NIU":c="683:NU:NIU:Niue"
Case "1-670","MP","MNP":c="1-670:MP:MNP:Northern Mariana Islands"
Case "850","KP","PRK":c="850:KP:PRK:North Korea"
Case "47","NO","NOR":c="47:NO:NOR:Norway"
Case "968","OM","OMN":c="968:OM:OMN:Oman"
Case "92","PK","PAK":c="92:PK:PAK:Pakistan"
Case "680","PW","PLW":c="680:PW:PLW:Palau"
Case "970","PS","PSE":c="970:PS:PSE:Palestine"
Case "507","PA","PAN":c="507:PA:PAN:Panama"
Case "675","PG","PNG":c="675:PG:PNG:Papua New Guinea"
Case "595","PY","PRY":c="595:PY:PRY:Paraguay"
Case "51","PE","PER":c="51:PE:PER:Peru"
Case "63","PH","PHL":c="63:PH:PHL:Philippines"
Case "64","PN","PCN":c="64:PN:PCN:Pitcairn"
Case "48","PL","POL":c="48:PL:POL:Poland"
Case "351","PT","PRT":c="351:PT:PRT:Portugal"
Case "1-787","1-939","PR","PRI":c="1-939:PR:PRI:Puerto Rico"
Case "974","QA","QAT":c="974:QA:QAT:Qatar"
Case "262","RE","REU":c="262:RE:REU:Reunion"
Case "40","RO","ROU":c="40:RO:ROU:Romania"
Case "7","RU","RUS":c="7:RU:RUS:Russia"
Case "250","RW","RWA":c="250:RW:RWA:Rwanda"
Case "590","BL","BLM":c="590:BL:BLM:Saint Barthelemy"
Case "685","WS","WSM":c="685:WS:WSM:Samoa"
Case "378","SM","SMR":c="378:SM:SMR:San Marino"
Case "239","ST","STP":c="239:ST:STP:Sao Tome and Principe"
Case "966","SA","SAU":c="966:SA:SAU:Saudi Arabia"
Case "221","SN","SEN":c="221:SN:SEN:Senegal"
Case "381","RS","SRB":c="381:RS:SRB:Serbia"
Case "248","SC","SYC":c="248:SC:SYC:Seychelles"
Case "232","SL","SLE":c="232:SL:SLE:Sierra Leone"
Case "65","SG","SGP":c="65:SG:SGP:Singapore"
Case "1-721","SX","SXM":c="1-721:SX:SXM:Sint Maarten"
Case "421","SK","SVK":c="421:SK:SVK:Slovakia"
Case "386","SI","SVN":c="386:SI:SVN:Slovenia"
Case "677","SB","SLB":c="677:SB:SLB:Solomon Islands"
Case "252","SO","SOM":c="252:SO:SOM:Somalia"
Case "27","ZA","ZAF":c="27:ZA:ZAF:South Africa"
Case "82","KR","KOR":c="82:KR:KOR:South Korea"
Case "211","SS","SSD":c="211:SS:SSD:South Sudan"
Case "34","ES","ESP":c="34:ES:ESP:Spain"
Case "94","LK","LKA":c="94:LK:LKA:Sri Lanka"
Case "290","SH","SHN":c="290:SH:SHN:Saint Helena"
Case "1-869","KN","KNA":c="1-869:KN:KNA:Saint Kitts and Nevis"
Case "1-758","LC","LCA":c="1-758:LC:LCA:Saint Lucia"
Case "590","MF","MAF":c="590:MF:MAF:Saint Martin"
Case "508","PM","SPM":c="508:PM:SPM:Saint Pierre and Miquelon"
Case "1-784","VC","VCT":c="1-784:VC:VCT:Saint Vincent and the Grenadines"
Case "249","SD","SDN":c="249:SD:SDN:Sudan"
Case "597","SR","SUR":c="597:SR:SUR:Suriname"
Case "47","SJ","SJM":c="47:SJ:SJM:Svalbard and Jan Mayen"
Case "268","SZ","SWZ":c="268:SZ:SWZ:Swaziland"
Case "46","SE","SWE":c="46:SE:SWE:Sweden"
Case "41","CH","CHE":c="41:CH:CHE:Switzerland"
Case "963","SY","SYR":c="963:SY:SYR:Syria"
Case "886","TW","TWN":c="886:TW:TWN:Taiwan"
Case "992","TJ","TJK":c="992:TJ:TJK:Tajikistan"
Case "255","TZ","TZA":c="255:TZ:TZA:Tanzania"
Case "66","TH","THA":c="66:TH:THA:Thailand"
Case "228","TG","TGO":c="228:TG:TGO:Togo"
Case "690","TK","TKL":c="690:TK:TKL:Tokelau"
Case "676","TO","TON":c="676:TO:TON:Tonga"
Case "1-868","TT","TTO":c="1-868:TT:TTO:Trinidad and Tobago"
Case "216","TN","TUN":c="216:TN:TUN:Tunisia"
Case "90","TR","TUR":c="90:TR:TUR:Turkey"
Case "993","TM","TKM":c="993:TM:TKM:Turkmenistan"
Case "1-649","TC","TCA":c="1-649:TC:TCA:Turks and Caicos Islands"
Case "688","TV","TUV":c="688:TV:TUV:Tuvalu"
Case "971","AE","ARE":c="971:AE:ARE:United Arab Emirates"
Case "256","UG","UGA":c="256:UG:UGA:Uganda"
Case "44","GB","UK","GBR":c="44:GB:GBR:United Kingdom"
Case "380","UA","UKR":c="380:UA:UKR:Ukraine"
Case "598","UY","URY":c="598:UY:URY:Uruguay"
Case "1","US","USA":c="1:US:USA:United States"
Case "998","UZ","UZB":c="998:UZ:UZB:Uzbekistan"
Case "678","VU","VUT":c="678:VU:VUT:Vanuatu"
Case "379","VA","VAT":c="379:VA:VAT:Vatican"
Case "58","VE","VEN":c="58:VE:VEN:Venezuela"
Case "84","VN","VNM":c="84:VN:VNM:Vietnam"
Case "1-340","VI","VIR":c="1-340:VI:VIR:U.S. Virgin Islands"
Case "681","WF","WLF":c="681:WF:WLF:Wallis and Futuna"
Case "212","EH","ESH":c="212:EH:ESH:Western Sahara"
Case "967","YE","YEM":c="967:YE:YEM:Yemen"
Case "260","ZM","ZMB":c="260:ZM:ZMB:Zambia"
Case "263","ZW","ZWE":c="263:ZW:ZWE:Zimbabwe"
End Select
If c<>"" Then
a=Split(c,":")
x("00")=a(0)
x("xx")=a(1)
x("xxx")=a(2)
x("name")=a(3)
End If
End If
Set CountryData=x
End Function
Function SiteFilterFor(ByVal nodes)
Dim k,o
k="#FWX_ENT:"&nodes&"__SITEFTR"
o=j(k)
If o="" Then
Dim node
For Each node In Split(nodes,",")
If o="" Then
o=EXE_Backward("FWX_ENT_"&node,"SITEFTR")
If o<>"" Then
o=Enclose(o,"(",")")
End If
End If
Next
If o="" Then o="!"
j(k)=o
End If
If o="!" Then o=""
SiteFilterFor=o
End Function
Dim CACHE__SiteFilter:CACHE__SiteFilter=""
Function SiteFilter()
If CACHE__SiteFilter<>"" Then
SiteFilter=CACHE__SiteFilter
Exit Function
End If
If IsObject(FWX)Then CACHE__SiteFilter=FWX.Setting("SiteFilter")
If CACHE__SiteFilter="" Then
CACHE__SiteFilter=CACHE__SiteFilter&"ISNULL(t.site,-9) IN "
CACHE__SiteFilter=CACHE__SiteFilter&"(%SITE,"
If IsObject(FWX)Then CACHE__SiteFilter=CACHE__SiteFilter&Enclose(FWX.Setting("family"),",","")
CACHE__SiteFilter=CACHE__SiteFilter&"-9)"
End If
SiteFilter=CACHE__SiteFilter
End Function
Dim CACHE__MerchantFilter:CACHE__MerchantFilter=""
Function MerchantFilter()
If CACHE__MerchantFilter<>"" Then
MerchantFilter=CACHE__MerchantFilter
Exit Function
End If
Dim iMerchant:iMerchant=VOZ(MyMerchant)
If iMerchant>0 Then
CACHE__MerchantFilter="(t.merchant="&iMerchant&")"
End If
MerchantFilter=TOT(CACHE__MerchantFilter,"(1=1)")
End Function
Dim CACHE__Site:CACHE__Site=0
Function MySite()
If CACHE__Site<>0 Then
MySite=CACHE__Site
Exit Function
End If
If IsObject(FWX)Then
If Not CACHE__Site<>0 And IsBack Then
If IsObject(FWX)Then CACHE__Site=VOZ(FWX.Setting("siteAdmin"))
End If
If Not CACHE__Site<>0 Then
If IsObject(FWX)Then CACHE__Site=VOZ(FWX.Setting("site"))
End If
End If
MySite=CACHE__Site
End Function
Dim CACHE__DomainCache:CACHE__DomainCache=""
Function MyDomainCache()
If CACHE__DomainCache<>"" Then
MyDomainCache=CACHE__DomainCache
Exit Function
End If
Dim domain
domain=LCase(rxReplace(MyDomain,"^((www|local|l\d|temp|test|dev|office|admin|oms|m|d\d|w\d|v\d)\.)?",""))&"/"
CACHE__DomainCache="/$$TEMP/"&domain&period
MyDomainCache=CACHE__DomainCache
End Function
Function MyMDomain()
MyMDomain="m."&MyDomain
End Function
Dim CACHE__WWWDomain:CACHE__WWWDomain=""
Function MyWWWDomain()
Dim s
s=MyDomain
If Left(s,4)="www." Then
CACHE__WWWDomain=s
ElseIf s="localhost" Or Left(s,6)="local." Or SERVER_IP="127.0.0.1" Then
CACHE__WWWDomain="www."&DOMAIN_NAME
ElseIf s="127.0.0.1" Then
CACHE__WWWDomain="www."&DOMAIN_NAME
ElseIf IsSubDomain Then
CACHE__WWWDomain=DOMAIN_NAME
Else
CACHE__WWWDomain="www."&s
End If
MyWWWDomain=CACHE__WWWDomain
End Function
Dim CACHE__TMPL_FOLDER
Function TMPL_FOLDER()
If CACHE__TMPL_FOLDER<>"" Then
TMPL_FOLDER=CACHE__TMPL_FOLDER
Exit Function
End If
Dim o:o=""
If IsObject(FWX)Then
If o="" Then o=FWX.Setting("tmpl.path")
If o="" Then o=FWX.Setting("tmpl.folder")
If o="" Then o=Application.Contents("tmpl.path")
If o="" Then o=Application.Contents("tmpl.folder")
If o="" Then o=FWX.Setting("ThemePath")
If o="" Then o=DATA_FOLDER(MyDomainForCode)
Else
If o="" Then o=Application.Contents("tmpl.path")
If o="" Then o=Application.Contents("tmpl.folder")
If o="" Then o=Application.Contents("ThemePath")
End If
If o="" Then o=FWX__TMPL__FOLDER
If o="" Then o=SITE_FOLDER
If IsObject(FWX)And CACHE__Domain<>"" Then CACHE__TMPL_FOLDER=o
TMPL_FOLDER=o
End Function
Dim CACHE__SITE_FOLDER
Function SITE_FOLDER()
If CACHE__SITE_FOLDER<>"" Then
SITE_FOLDER=CACHE__SITE_FOLDER
Exit Function
End If
If CACHE__SITE_FOLDER="" Then
CACHE__SITE_FOLDER=DATA_FOLDER(MyDomain)
End If
SITE_FOLDER=CACHE__SITE_FOLDER
End Function
Function DATA_FOLDER(ByVal s)
If s<>"" Then
s=Replace(s,".","-")
s="/$"&s&"/"
End If
DATA_FOLDER=s
End Function
Dim CACHE__Hostname:CACHE__Hostname=""
Function MyHostname()
If CACHE__Hostname="" Then
CACHE__Hostname=rxReplace(SERVER_NAME,"^www\.","")
End If
MyHostname=CACHE__Hostname
End Function
Function MyDomainID()
MyDomainID=UCase(Left(rxReplace(TOT(MyDomain,SERVER_NAME),"\W+",""),6))
End Function
Function MyDomainServer()
MyDomainServer=MyServerName
End Function
Function MyServerName()
MyServerName=rxReplace(SERVER_NAME,"^(www|local|temp|test|dev|office|admin|d\d|svn|wap|m|w\d|v\d)\.","")
End Function
Dim CACHE__Domain_ByServerName:CACHE__Domain_ByServerName=""
Dim CACHE__Domain:CACHE__Domain=""
Function MyDomain()
If CACHE__Domain<>"" Then
MyDomain=CACHE__Domain
Exit Function
End If
Dim o:o=""
If IsObject(FWX)Then
If o="" Then o=FWX.Setting("domain")
End If
If DOMAIN_NAME<>"localhost" Then
If o="" Then o=rxReplace(DOMAIN_NAME,"^www\.","")
End If
o=rxReplace(rxReplace(o,"^www\.",""),"^[\s\n\r]+|[\s\n\r]+$","")
If IsObject(FWX)And o<>"" Then CACHE__Domain=o
MyDomain=o
End Function
Dim CACHE__DomainForData:CACHE__DomainForData=""
Function MyDomainForData()
If CACHE__DomainForData<>"" Then
MyDomainForData=CACHE__DomainForData
Exit Function
End If
Dim o:o=""
If IsObject(FWX)Then
If o="" Then o=FWX.Setting("domain.data."&DOMAIN_NAME)
If o="" Then o=FWX.Setting("domain.data")
If o="" Then o=DOMAIN_NAME_FOR_DATA
End If
If o="" Then o=MyDomain
If IsObject(FWX)And CACHE__Domain<>"" Then CACHE__DomainForData=o
MyDomainForData=o
End Function
Dim CACHE__DomainForCode:CACHE__DomainForCode=""
Function MyDomainForCode()
If CACHE__DomainForCode<>"" Then
MyDomainForCode=CACHE__DomainForCode
Exit Function
End If
Dim o:o=""
If IsObject(FWX)Then
If o="" Then o=FWX.Setting("domain.code."&DOMAIN_NAME)
If o="" Then o=FWX.Setting("domain.code")
If o="" Then o=Application.Contents("domain.code")
If o="" Then o=DOMAIN_NAME_FOR_CODE
End If
If o="" Then o=MyDomain
If IsObject(FWX)And CACHE__Domain<>"" Then CACHE__DomainForCode=o
MyDomainForCode=o
End Function
Dim CACHE__Passive:CACHE__Passive=""
Function IsPassive()
If CACHE__Passive<>"" Then
IsPassive=CBool(CACHE__Passive)
Exit Function
End If
CACHE__Passive=True
DBG "IsPassive? IsAjax="&IsAjax
DBG "IsPassive? IsPost="&IsPost
DBG "IsPassive? IsRobot="&IsRobot
DBG "IsPassive? IsStaffPC="&IsStaffPC
DBG "IsPassive? IsLoggedIn="&IsLoggedIn
DBG "IsPassive? IsFrontpage="&IsTrue(j("-FP"))
DBG "IsPassive? REFERRER="&REFERRER
DBG "IsPassive? IsObject(quote)="&IsObject(quote)
DBG "IsPassive? MyQuote="&MyQuote
If IsPost Then
CACHE__Passive=False
ElseIf IsRobot Then
CACHE__Passive=True
ElseIf IsStaffPC Then
CACHE__Passive=False
ElseIf IsLoggedIn Then
CACHE__Passive=False
ElseIf IsTrue(j("-FP"))Then
CACHE__Passive=True
ElseIf rxTest(REFERRER,"google|yahoo|bing|mail")Then
CACHE__Passive=True
Else
If MyQuote>0 Then
If IsObject(quote)Then
If VOZ(quote.Field("items"))>0 Then
CACHE__Passive=False
End If
End If
End If
End If
IsPassive=CBool(CACHE__Passive)
End Function
Function UseSuperCache()
UseSuperCache=CBool(IsPassive And Query.EXT="")' And Query.String=""
End Function
Dim CACHE__Country:CACHE__Country=""
Function MyCountry()
If CACHE__Country="" Then
CACHE__Country=MySession("country")
If CACHE__Country="" Then
If CACHE__Country="" Then
CACHE__Country=Request.ServerVariables("HTTP_CF-IPCOUNTRY")
If CACHE__Country="ZZ" Then CACHE__Country=""
End If
If CACHE__Country="" Then
CACHE__Country=GeoIPLocCode(VISITOR_IP)
If CACHE__Country="ZZ" Then CACHE__Country=""
End If
If CACHE__Country="" Then
CACHE__Country=GeoIPLocCode(SERVER_IP)
If CACHE__Country="ZZ" Then CACHE__Country=""
End If
MySession("country")=CACHE__Country
End If
CACHE__Country=UCase(CACHE__Country)
End If
If CACHE__Country="GB" Then
CACHE__Country="UK"
End If
MyCountry=Trim(CACHE__Country)
End Function
Function MyCountryCode()
MyCountryCode=MyCountry
End Function
Dim CACHE__Currency:CACHE__Currency=""
Function MyCurrency()
If CACHE__Currency<>"" Then
MyCurrency=Trim(CACHE__Currency)
Exit Function
End If
If CACHE__Currency="" Then
If IsObject(MySession)Then CACHE__Currency=MySession("cur")
End If
If IsObject(quote)Then
If IsObject(Query)Then
If Query.FQ("Quote.New,Quote.Close")="" Then
CACHE__Currency=TextOrSpace(quote.Field("cur"))
End If
End If
End If
If CACHE__Currency="" Then
CACHE__Currency=j("cur")
End If
If CACHE__Currency="" Then
CACHE__Currency=Application.Contents("cur")
End If
MySession("cur")=Trim(CACHE__Currency)
MyCurrency=Trim(CACHE__Currency)
End Function
Dim CACHE__Username:CACHE__Username=""
Function MyUsername()
If CACHE__Username="" Then
CACHE__Username=TextOrSpace(rxSelect(MyEmail,"^[^@]+"))
End If
MyUsername=Trim(CACHE__Username)
End Function
Dim CACHE__Email:CACHE__Email=""
Function MyEmail()
If CACHE__Email<>"" Then
MyEmail=Trim(CACHE__Email)
Exit Function
End If
If IsObject(visitor)Then
CACHE__Email=TextOrSpace(visitor.Field("email"))
If CACHE__Email="" Then CACHE__Email=" "
End If
If CACHE__Email="" Then
If IsObject(MySession)Then CACHE__Email=MySession("email")
End If
MyEmail=Trim(CACHE__Email)
End Function
Dim CACHE__Name:CACHE__Name=""
Function MyName()
If CACHE__Name<>"" Then
MyName=Trim(CACHE__Name)
Exit Function
End If
If IsObject(visitor)Then
CACHE__Name=TextOrSpace(visitor.Field("name"))
If CACHE__Name="" Then CACHE__Name=" "
End If
If CACHE__Name="" Then
If IsObject(MySession)Then CACHE__Name=MySession("name")
End If
MyName=Trim(CACHE__Name)
End Function
Dim CACHE__Credentials:CACHE__Credentials=""
Function MyCredentials()
If CACHE__Credentials<>"" Then
MyCredentials=Trim(CACHE__Credentials)
Exit Function
End If
If CACHE__Credentials="" Then
If IsObject(MySession)Then CACHE__Credentials=MySession("credentials")
End If
If CACHE__Credentials="" Then
If IsObject(Session)Then CACHE__Credentials=Session.Contents("credentials")
End If
If IsObject(visitor)Then
If Not visitor.Started Then visitor.Start()
CACHE__Credentials=TextOrSpace(visitor.Field("credentials"))
If CACHE__Credentials="" Then CACHE__Credentials=" "
End If
CACHE__Credentials=rxReplace(CACHE__Credentials,"\W+",",")
CACHE__Credentials=rxReplace(CACHE__Credentials,"^,|,$","")
MyCredentials=Trim(CACHE__Credentials)
End Function
Dim CACHE__VisitorCookieID:CACHE__VisitorCookieID=""
Function MyVisitorCookieID()
If CACHE__VisitorCookieID="" Then
CACHE__VisitorCookieID="V"&Encrypt(IIf(IsBack,"BACKER","FRONTE")&"")
End If
MyVisitorCookieID=Trim(CACHE__VisitorCookieID)
End Function
Dim CACHE__Visitor:CACHE__Visitor=""
Function MyVisitor()
If CACHE__Visitor<>"" Then
MyVisitor=Trim(CACHE__Visitor)
Exit Function
End If
If IsObject(visitor)Then CACHE__Visitor=visitor.ID
If Not(CACHE__Visitor<>"")And IsObject(MyCookies)Then CACHE__Visitor=MyCookies.BlockBased(MyVisitorCookieID)
If Not(CACHE__Visitor<>"")And IsObject(MySession)Then CACHE__Visitor=MySession(MyVisitorCookieID)
MyVisitor=Trim(CACHE__Visitor)
End Function
Function MyStateKey():MyStateKey="S"&MyVisitor:End Function
Function MyLPingKey():MyLPingKey="P"&MyVisitor:End Function
Function MyState()
If j("MyState")="" Then
j("MyState")=MySession(MyStateKey)
End If
MyState=j("MyState")
End Function
Function MyLPing()
If j("MyLPing")="" Then
j("MyLPing")=MySession(MyLPingKey)
End If
MyLPing=j("MyLPing")
End Function
Function MyStateAge()
MyStateAge=DateDiff("s",DON(MyLPing),Now())
End Function
Function MyStateExpiry()
MyStateExpiry=FWX.Setting("visitor.state.expiry")
End Function
Const FWX__visitor_state_expiry=600
Function MyStateAutoStart()
MyStateAutoStart=IsTrue(FWX.Setting("visitor.state.auto"))
End Function
Const FWX__visitor_state_auto="n"
Dim CACHE__Quote:CACHE__Quote=""
Function MyQuote()
If CACHE__Quote<>"" Then
MyQuote=CLng("0"&CACHE__Quote)
Exit Function
End If
If IsFront Then
If IsObject(visitor)Then CACHE__Quote=visitor.Field("quote")
If VOZ("0"&CACHE__Quote)=0 Then'CACHE__Quote="" Then
If IsObject(MySession)Then CACHE__Quote=MySession("quote")
End If
If VOZ("0"&CACHE__Quote)=0 Then'CACHE__Quote="" Then
If IsObject(MySession)Then CACHE__Quote=MyCookies("quote")
End If
Else
If VOZ("0"&CACHE__Quote)=0 Then'CACHE__Quote="" Then
If IsObject(Query)Then CACHE__Quote=Query.FQ("quote")
End If
End If
If TOE(CACHE__Quote)<>"" Then
If VOZ("0"&CACHE__Quote)=0 Then
CACHE__Quote=""
End If
End If
MyQuote=CLng("0"&CACHE__Quote)
End Function
Dim CACHE__Corporate:CACHE__Corporate=""
Function MyCorporate()
Dim rShopperFlags
If CACHE__Corporate<>"" Then
MyCorporate=CBool(CACHE__Corporate)
Exit Function
End If
If CACHE__Corporate="" And IsBack Then
If IsObject(quot)Then
iQ=VOZ(Query.FQ("quote"))
If iQ>0 Then
Set rShopperFlags=quot.Carts.Run("SELECT corporate,vip/*ShopperFlags*/ FROM %TABLEU WHERE %KEY="&iQ)
If Exists(rShopperFlags)Then CACHE__Corporate=rShopperFlags("corporate")
If Exists(rShopperFlags)Then CACHE__VIP=rShopperFlags("vip")
KillRS rShopperFlags
End If
End If
Else
If CACHE__Corporate="" And IsFront Then
If IsObject(visitor)Then
Set rShopperFlags=visitor.Run("SELECT corporate,vip/*ShopperFlags*/ FROM %TABLEU WHERE %KEY=%ID")
If Exists(rShopperFlags)Then CACHE__Corporate=rShopperFlags("corporate")
If Exists(rShopperFlags)Then CACHE__VIP=rShopperFlags("vip")
KillRS rShopperFlags
End If
End If
If CACHE__Corporate="" And IsFront Then
If IsObject(quote)Then CACHE__Corporate=quote.Field("corporate")
If CACHE__Corporate="" And IsFront Then
If IsObject(MySession)Then CACHE__Corporate=MySession("corporate")
End If
End If
End If
If CACHE__Corporate<>"" Then
MyCorporate=CBool(CACHE__Corporate)
Else
MyCorporate=False
End If
End Function
Dim CACHE__VIP:CACHE__VIP=""
Function MyVIP()
Dim rShopperFlags
If CACHE__VIP<>"" Then
MyVIP=CBool(CACHE__VIP)
Exit Function
End If
If CACHE__VIP="" And IsBack Then
If IsObject(quot)Then
iQ=VOZ(Query.FQ("quote"))
If iQ>0 Then
Set rShopperFlags=quot.Carts.Run("SELECT corporate,vip/*ShopperFlags*/ FROM %TABLEU WHERE %KEY="&iQ)
If Exists(rShopperFlags)Then CACHE__Corporate=rShopperFlags("corporate")
If Exists(rShopperFlags)Then CACHE__VIP=rShopperFlags("vip")
KillRS rShopperFlags
End If
End If
Else
If CACHE__VIP="" And IsFront Then
If IsObject(visitor)Then
Set rShopperFlags=visitor.Run("SELECT corporate,vip/*ShopperFlags*/ FROM %TABLEU WHERE %KEY=%ID")
If Exists(rShopperFlags)Then CACHE__Corporate=rShopperFlags("corporate")
If Exists(rShopperFlags)Then CACHE__VIP=rShopperFlags("vip")
KillRS rShopperFlags
End If
End If
If CACHE__VIP="" Then
If IsObject(quote)Then CACHE__VIP=quote.Field("vip")
If CACHE__VIP="" And IsFront Then
If IsObject(MySession)Then CACHE__VIP=MySession("vip")
End If
End If
End If
If CACHE__VIP<>"" Then
MyVIP=CBool(CACHE__VIP)
Else
MyVIP=False
End If
End Function
Dim CACHE__Trade:CACHE__Trade=""
Function MyTrade()
If CACHE__Trade<>"" Then
MyTrade=CBool(CACHE__Trade)
Exit Function
End If
If IsObject(visitor)Then CACHE__Trade=visitor.Field("trade")
If CACHE__Trade="" Then
If IsObject(quote)Then CACHE__Trade=quote.Field("trade")
If CACHE__Trade="" Then
If IsObject(MySession)Then CACHE__Trade=MySession("trade")
End If
End If
If CACHE__Trade<>"" Then
MyTrade=CBool(CACHE__Trade)
Else
MyTrade=False
End If
End Function
Dim CACHE__Merchant:CACHE__Merchant=""
Function MyMerchant()
If CACHE__Merchant<>"" Then
MyMerchant=CACHE__Merchant
Exit Function
End If
If CACHE__Merchant="" Then
If IsObject(MySession)Then
bTrade=IsTrue(MySession("trade"))
End If
If Not bTrade Then
bTrade=IsTrue(MyTrade)
End If
If bTrade Then
CACHE__Merchant=MyAccount
Else
CACHE__Merchant=0
End If
End If
MyMerchant=CLng("0"&CACHE__Merchant)
End Function
Dim CACHE__Account:CACHE__Account=""
Function MyAccount()
If CACHE__Account<>"" Then
MyAccount=CLng("0"&CACHE__Account)
Exit Function
End If
If IsObject(visitor)Then CACHE__Account=VOZ(visitor.Field("account")&"")
If CACHE__Account="" Then
If IsObject(MySession)Then CACHE__Account=MySession("account")
End If
MyAccount=CLng("0"&CACHE__Account)
End Function
Dim CACHE__quoteAccount:CACHE__quoteAccount=""
Function MyQuoteAccount()
If CACHE__quoteAccount<>"" Then
MyQuoteAccount=CLng("0"&CACHE__quoteAccount)
Exit Function
End If
If IsObject(visitor)Then CACHE__quoteAccount=visitor.Field("quoteAccount")
If CACHE__quoteAccount="" Then
If IsObject(quote)Then CACHE__quoteAccount=VOZ(quote.Field("account")&"")
If CACHE__quoteAccount="" Then
If IsObject(MySession)Then CACHE__quoteAccount=MySession("quoteAccount")
End If
End If
MyQuoteAccount=CLng("0"&CACHE__quoteAccount)
End Function
Dim CACHE__quoteStockist:CACHE__quoteStockist=""
Function MyQuoteStockist()
If CACHE__quoteStockist_overrideChecked="" Then
If IsObject(Query)Then
CACHE__quoteStockist_overrideChecked="y"
If Query.FQ("--stockist")<>"" Then
CACHE__quoteStockist=Query.FQ("--stockist")
End If
End If
End If
If CACHE__quoteStockist<>"" Then
MyQuoteStockist=CLng("0"&CACHE__quoteStockist)
Exit Function
End If
If CACHE__QuoteStockist="" Then
If IsObject(Query)Then CACHE__QuoteStockist=Query.FQ("--stockist")
End If
If CACHE__QuoteStockist="" Then
If IsObject(FWX)Then CACHE__QuoteStockist=FWX.Setting("quot.quotes.stockist")
End If
If CACHE__QuoteStockist="" Then
If IsObject(FWX)Then CACHE__QuoteStockist=FWX.Setting("stockist")
End If
If IsObject(visitor)Then CACHE__quoteStockist=TOE(visitor.Field("quoteStockist")&"")
If CACHE__quoteStockist="" Then
If IsObject(quote)Then CACHE__quoteStockist=TOE(quote.Field("stockist")&"")
If CACHE__quoteStockist="" Then
If IsObject(MySession)Then CACHE__quoteStockist=MySession("quoteStockist")
End If
End If
MyQuoteStockist=CLng("0"&CACHE__quoteStockist)
End Function
Dim CACHE__quoteMerchant:CACHE__quoteMerchant=""
Function MyQuoteMerchant()
If TOE(CACHE__quoteMerchant)<>"" Then
MyQuoteMerchant=VOZ("0"&CACHE__quoteMerchant)
Exit Function
End If
If CACHE__QuoteMerchant="" Then
If IsObject(Query)Then CACHE__QuoteMerchant=Query.FQ("--merchant")
End If
If CACHE__QuoteMerchant="" Then
If IsObject(FWX)Then CACHE__QuoteMerchant=FWX.Setting("quot.quotes.merchant")
End If
If CACHE__QuoteMerchant="" Then
If IsObject(FWX)Then CACHE__QuoteMerchant=FWX.Setting("merchant")
End If
If CACHE__QuoteMerchant="" Then
If IsObject(visitor)Then CACHE__quoteMerchant=visitor.Field("quoteMerchant")
End If
If CACHE__quoteMerchant="" Then
If IsObject(quote)Then CACHE__quoteMerchant=VOZ(quote.Field("merchant")&"")
If CACHE__quoteMerchant="" Then
If IsObject(MySession)Then CACHE__quoteMerchant=MySession("quoteMerchant")
End If
End If
MyQuoteMerchant=CLng("0"&CACHE__quoteMerchant)
End Function
Function CanAccess(ByVal sReq)
CanAccess=HasAccess(MyCredentials,sReq)
End Function
Dim dAccessCache:Set dAccessCache=Dictionary()
Function HasAccess(ByVal sCred,ByVal sReq)
sCred=LCase(Trim(TOE(sCred)))
sCred=Replace(Trim(rxReplace(sCred,"\W+"," "))," ",",")
If sCred="" Then
HasAccess=False:Exit Function
Else
sCred="["&Join(Split(sCred,","),"],[")&"]"
End If
sReq=LCase(Trim(TOE(sReq)))
sReq=Replace(Trim(rxReplace(sReq,"\W+"," "))," ",",")
If sReq="" Or sReq="[public]" Then
HasAccess=True:Exit Function
Else
sReq="["&Join(Split(sReq,","),"],[")&"]"
End If
sReq=rxReplace(sReq,"[\W\s\t]+",",")
sReq=rxReplace(sReq,"^,|,$","")
dAccessCache(sReq)=dAccessCache(sReq)&""
If dAccessCache(sReq)<>"" Then
HasAccess=CBool(dAccessCache(sReq)):Exit Function
End If
sCred=LCase(Trim(TOE(Replace(Replace(sCred,"[",",["),"]","],"))))
sCred=rxReplace(sCred,"[\W\s\t]+",",")
sCred=rxReplace(sCred,"^,|,$","")
If sCred<>"" Then
Dim bResult,cs,c,lResult
bResult=False
cs=Split(sCred,",")
For Each c In cs
c=Trim(c&"")
If c<>"" Then
lResult=False
lResult=lResult Or CBool(InStr(1,sReq,c)>0)
lResult=lResult Or CBool(AccessLevel(c)>AccessLevel(sReq))
bResult=bResult Or lResult
End If
Next
dAccessCache(sReq)=bResult
End If
HasAccess=bResult
End Function
Function IsAccess(ByVal sReq)
IsAccess=CBool(MyAccess&""=sReq&"")
End Function
Function IsAllowed(ByVal sReq)
IsAllowed=HasAccess(MyCredentials,LCase(sReq&""))
End Function
Function IsLoggedIn()
IsLoggedIn=CBool(GTZ(MyAccount)Or IsSet(MyEmail)Or IsTrue(j("@AUTHORIZED")))'(MyEmail&""<>"" And IsBack))
End Function
Function IsTrade()
IsTrade=CBool(MyTrade Or MyMerchant>0)
End Function
Function IsQuoteTrade()
IsQuoteTrade=CBool(MyQuoteMerchant>0)
End Function
Function IsStaffPC()
Dim b,k
b=False
k="IsStaff.Flag"
If CACHE__Credentials<>"" And CACHE__Email<>"" And CACHE__Account<>"" Then
If Not b Then
b=CBool(IsStaff)
End If
End If
If Not b Then
If IsObject(MyCookies)Then
b=CBool(MyCookies.Global(k)<>"")
End If
End If
If Not b Then
If IsObject(Request)Then
If IsObject(Request.Cookies)Then
b=CBool(Request.Cookies(k)<>"")
End If
End If
End If
If Not b Then
If IsObject(MySession)Then
b=CBool(MySession(k)<>"")
End If
End If
If Not b Then
If IsObject(Session)Then
If IsObject(Session.Contents)Then
b=CBool(Session.Contents(k)<>"")
End If
End If
End If
IsStaffPC=b
End Function
Function AccessLevel(ByVal c)
c=TOE(c)
If dAccessCache("L"&c)="" Then
dAccessCache("L"&c)=0
For Each x In Split(c,",")
dAccessCache("L"&x)=VOZ(EXE("FWX_Access_Level_"&x))
If dAccessCache("L"&x)>dAccessCache("L"&c)Then
dAccessCache("L"&c)=dAccessCache("L"&x)
End If
Next
End If
AccessLevel=dAccessCache("L"&c)
End Function
Const FWX_Access_Level_=1
Const FWX_Access_Level_public=1
Const FWX_Access_Level_member=5
Const FWX_Access_Level_client=5
Const FWX_Access_Level_customer=5
Const FWX_Access_Level_driver=7
Const FWX_Access_Level_controller=50
Const FWX_Access_Level_clerk=100
Const FWX_Access_Level_staff=200
Const FWX_Access_Level_stock=222
Const FWX_Access_Level_supervisor=300
Const FWX_Access_Level_stores=350
Const FWX_Access_Level_editor=333
Const FWX_Access_Level_marketing=350
Const FWX_Access_Level_accountant=400
Const FWX_Access_Level_manager=500
Const FWX_Access_Level_director=700
Const FWX_Access_Level_ceo=800
Const FWX_Access_Level_owner=800
Const FWX_Access_Level_admin=900
Const FWX_Access_Level_ogeid=999
Const FWX_Access_Level_master=9999
Function IsMember()
IsMember=IsAllowed("[member]"&fwxMasterAccess)
End Function
Function IsStaff()
IsStaff=IsAllowed("[admin][ceo][director][manager][editor][clerk][staff]"&fwxMasterAccess)
End Function
Function IsClerk()
IsClerk=IsAllowed("[admin][ceo][director][manager][clerk][staff]"&fwxMasterAccess)
End Function
Function IsEditor()
IsEditor=IsAllowed("[admin][ceo][director][manager][editor]"&fwxMasterAccess)
End Function
Function IsMarketing()
IsMarketing=IsAllowed("[admin][ceo][director][manager][editor][marketing]"&fwxMasterAccess)
End Function
Function IsSupervisor()
IsSupervisor=IsAllowed("[admin][ceo][director][manager][supervisor][accountant]"&fwxMasterAccess)
End Function
Function IsStock()
IsStock=IsAllowed("[admin][ceo][director][manager][supervisor][stock]"&fwxMasterAccess)
End Function
Function IsController()
IsController=IsAllowed("[admin][ceo][director][manager][supervisor][controller]"&fwxMasterAccess)
End Function
Function IsAccountant()
IsAccountant=IsAllowed("[admin][ceo][director][manager][accountant]"&fwxMasterAccess)
End Function
Function IsManager()
IsManager=IsAllowed("[admin][ceo][director][manager][admin]"&fwxMasterAccess)
End Function
Function IsAdmin()
IsAdmin=IsAllowed("[admin]"&fwxMasterAccess)
End Function
Function IsMaster()
IsMaster=IsAllowed(fwxMasterAccess)
End Function
Function DefaultRequiredAccess()
DefaultRequiredAccess=""
If IsBack Then
DefaultRequiredAccess="[manager][admin][ceo][editor][clerk][staff]"
End If
End Function
Dim ThisResource
ThisResource=Request.ServerVariables("PATH_INFO")
ThisResource=Left(ThisResource,InStrRev(ThisResource,"/"))
ThisResource=Replace(ThisResource,"/fwx/back/","/")
Dim MyPermissions
Sub LoadSecuritySettings()
MyPermissions="RWED"
End Sub
Function CanI(ByVal sPerm)
If MyPermissions="" Then
Call LoadSecuritySettings()
End If
CanI=CBool(InStr(1,MyPermissions,sPerm)>0)
End Function
Function CanIError(ByVal sPerm)
DBG.FatalError "Access Denied (CanI)"
End Function
Function CanIRead()
Dim bCanI:bCanI=CanI("R")
If IsBack And Not bCanI Then Call CanIError("R")
CanIRead=bCanI
End Function
Function CanIWrite()
Dim bCanI:bCanI=CanI("W")
If IsBack And Not bCanI Then Call CanIError("W")
CanIWrite=bCanI
End Function
Function CanIEdit()
Dim bCanI:bCanI=CanI("E")
If IsBack And Not bCanI Then Call CanIError("E")
CanIEdit=bCanI
End Function
Function CanIDelete()
Dim bCanI:bCanI=CanI("D")
If IsBack And Not bCanI Then Call CanIError("D")
CanIDelete=bCanI
End Function
j("debug")=CBool(IsTest Or DBG.Show)
Function FWX_HTTP__VerbHandler()
If HTTP_VerbHandler="" Then
Call DEFAULT_HTTP_VerbHandler()
End If
End Function
Call FWX_HTTP__VerbHandler()
Function DEFAULT_HTTP_VerbHandler()
Select Case UCase(REQUEST_METHOD)
Case "PUT":Call FWX_HTTP__PUT
Case "GET":Call FWX_HTTP__GET
Case "POST":Call FWX_HTTP__POST
Case "DELETE":Call FWX_HTTP__DELETE
Case "OPTIONS":Call FWX_HTTP__OPTIONS
Case "HEAD":Call FWX_HTTP__HEAD
Case "":Call FWX_HTTP__EMPTY
Case Else:Call FWX_HTTP__UNKONWN
End Select
If REQUEST_METHOD="OPTIONS" Or REQUEST_METHOD="HEAD" Then
Response.End()
Else
If REQUEST_ORIGIN<>SERVER_NAME And REQUEST_ORIGIN<>"" Then
Call FWX_HTTP__CheckOrigin()
End If
End If
End Function
Function FWX_HTTP__GET()
If HTTP_VerbHandler_GET="" Then
Call DEFAULT_HTTP_VerbHandler_GET()
End If
End Function
Function FWX_HTTP__POST()
If HTTP_VerbHandler_POST="" Then
Call DEFAULT_HTTP_VerbHandler_POST()
End If
End Function
Function FWX_HTTP__PUT()
If HTTP_VerbHandler_PUT="" Then
Call DEFAULT_HTTP_VerbHandler_PUT()
End If
End Function
Function FWX_HTTP__DELETE()
If HTTP_VerbHandler_DELETE="" Then
Call DEFAULT_HTTP_VerbHandler_DELETE()
End If
End Function
Function FWX_HTTP__OPTIONS()
If HTTP_VerbHandler_OPTIONS="" Then
Call DEFAULT_HTTP_VerbHandler_OPTIONS()
End If
End Function
Function FWX_HTTP__HEAD()
If HTTP_VerbHandler_HEAD="" Then
Call DEFAULT_HTTP_VerbHandler_HEAD()
End If
End Function
Function FWX_HTTP__EMPTY()
If HTTP_VerbHandler_EMPTY="" Then
Call DEFAULT_HTTP_VerbHandler_EMPTY()
End If
End Function
Function FWX_HTTP__UNKONWN()
If HTTP_VerbHandler_UNKONWN="" Then
Call DEFAULT_HTTP_VerbHandler_UNKONWN()
End If
End Function
Function DEFAULT_HTTP_VerbHandler_GET()
End Function
Function DEFAULT_HTTP_VerbHandler_POST()
End Function
Function DEFAULT_HTTP_VerbHandler_PUT()
End Function
Function DEFAULT_HTTP_VerbHandler_DELETE()
End Function
Function DEFAULT_HTTP_VerbHandler_OPTIONS()
Call FWX_HTTP__CheckOrigin()
End Function
Function DEFAULT_HTTP_VerbHandler_HEAD()
End Function
Function DEFAULT_HTTP_VerbHandler_EMPTY()
End Function
Function DEFAULT_HTTP_VerbHandler_UNKNOWN()
End Function
Function FWX_HTTP__AllowOrigin()
If HTTP_Headers_AllowOrigin="" Then
Call DEFAULT_HTTP_Headers_AllowOrigin()
End If
End Function
Dim DEFAULT_HTTP_Headers_AllowOrigin__EXECUTED
DEFAULT_HTTP_Headers_AllowOrigin__EXECUTED=False
Function DEFAULT_HTTP_Headers_AllowOrigin()
If DEFAULT_HTTP_Headers_AllowOrigin__EXECUTED Then
Else
DEFAULT_HTTP_Headers_AllowOrigin__EXECUTED=True
Response.AddHeader "Access-Control-Allow-Origin",TOT(FWX__access_control_allow_origin_header,TOT(REQUEST_ORIGIN,"*"))
Response.AddHeader "Access-Control-Allow-Headers",TOT(FWX__access_control_allow_headers_header,"Origin, X-Requested-With, X-Method, auth, Accept, Accept-Encoding, Content-Type, Content-Encoding")
Response.AddHeader "Access-Control-Allow-Methods",TOT(FWX__access_control_allow_methods_header,"PUT, GET, POST, DELETE, OPTIONS, JSON")
Response.AddHeader "Access-Control-Expose-Headers",TOT(FWX__access_control_expose_headers_header,"*")'"token")
Response.AddHeader "Access-Control-Allow-Credentials",TOT(FWX__access_control_allow_credentials_header,"true")
End If
End Function
Function FWX_HTTP__CheckOrigin()
If HTTP_Headers_CheckOrigin="" Then
Call DEFAULT_HTTP_Headers_CheckOrigin()
End If
End Function
Dim DEFAULT_HTTP_Headers_CheckOrigin__EXECUTED
DEFAULT_HTTP_Headers_CheckOrigin__EXECUTED=False
Function DEFAULT_HTTP_Headers_CheckOrigin()
If DEFAULT_HTTP_Headers_CheckOrigin__EXECUTED Then
Else
DEFAULT_HTTP_Headers_CheckOrigin__EXECUTED=True
Response.AddHeader "X-Access-Control-Handler","Fyneworks"
Dim tldless1,tldless2,allowed,result
source=TOT(REQUEST_ORIGIN,REFERRER)
result=False
allowed=FWX__access_control_allow_origin
domain1=GetDomain(source)
domain2=GetDomain(MyDomain)
tldless1=GetTLDLess(source)
tldless2=GetTLDLess(MyDomain)
If rxTest(domain1,"^(\w+\.)?icenium\.com$")Then
Call FWX_HTTP__AllowOrigin()
Exit Function
result=True
Else
If tldless1=tldless2 Or domain1=domain2 Or allowed="*" Or IsInStrList(REQUEST_ORIGIN,allowed)Then
result=True
End If
End If
If result Then
Call FWX_HTTP__AllowOrigin()
Else
Response.AddHeader "X-Access-Control","Denied"
End If
End If
End Function
Function Cur(ByVal a)
Dim s,m,c
m=Money(a)
If IsObject(j)Then
c=j("cur")
If c<>"" Then s=j("cur-"&c&"-symbol")
End If
If s="" Then
If c="" Then c=Application.Contents("cur")
If c="" Then c="GBP"
s=j("cur-"&c&"-symbol")
If s="" Then
s=TOT(EXE("FWX_Cur_"&c),"&pound;")
If IsObject(j)Then
j("cur-"&c&"-symbol")=s
End If
End If
End If
Cur=s&m
End Function
Function Money(ByVal s)
s=VOZ(s)
Money=FormatNumber(s,2)
End Function
Function DoEmail(ByVal rec)
output=""
If IsObject(rec)Then
If Not ExistsOrKill(rec)Then Exit Function
str=ColItem(rec,"email")
Else
str=rec
End If
If Not IsEmail(str)Then Exit Function
output=str
DoEmail=output
End Function
Function DoPhone(ByVal str)
output=""
str=Trim(str&"")
If str="" Then
Exit Function
End If
sFirstChar=Left(str,1)&""
If sFirstChar="(" Then
output=str
Else
str=Replace(str," ","")
If sFirstChar="+" Then
output=str
ElseIf Not IsPhone(str)Then
output=str
Else
If sFirstChar<>"0" Then str="0"&str
If Len(str)=11 Then
If 1=0 Then
ElseIf Left(str,2)="07" Then
output=Mid(str,1,5)&" "&Mid(str,6,3)&" "&Mid(str,9,3)
ElseIf Left(str,2)="08" Then
output=Mid(str,1,4)&" "&Mid(str,5,3)&" "&Mid(str,8,4)&""
ElseIf Left(str,2)="02" Then
output=Mid(str,1,4)&" "&Mid(str,5,3)&" "&Mid(str,8,4)
Else
output=Mid(str,1,5)&" "&Mid(str,6,3)&" "&Mid(str,9,3)
End If
If Len(str)>=12 Then
output=output&" "&Mid(str,12)
End If
Else
output=str
End If
End If
End If
DoPhone=output
End Function
Function DoPostcode(ByVal s)
On Error Resume Next
Dim sO,iSP
s=UCase(s&"")
sO=s
iSP=InStr(1,s," ")
If iSP>0 Then
sO=Left(s,iSP-1)
Else
If Len(s)>4 Then
sO=Left(s,Len(s)-3)
End If
End If
If sO<>"" Then
If Not InStr(1,"0123456789",Right(sO,1))>0 Then
sO=Left(sO,Len(sO)-1)
End If
End If
DoPostcode=sO
End Function
Function DoDate(ByVal s)
s=TOE(s)
If s="" Then Exit Function
If IsDate(s)And Len(s)>6 Then
Else
Select Case LCase(s)
Case "now":
s=Now()
Case "season":
s=""
If IsObject(shop)Then s=shop.SeasonStart()
If s="" Then s=Application.Contents("archive.date")
If s="" Then s="1-apr-"&Year(Now())
s=MakeDate(Day(s),Month(s),Year(Now())+IIf(Month(Now())>3,0,-1))
Case "yesterday","yday":
s=Date()-1
Case "today","2day":
s=Date()
Case "tomorrow","tomorro","2moro":
s=Date()+1
Case Else:
If 1=0 Then
ElseIf VOZ(s)<>0 Then
ElseIf rxTest(s,"^\d{1,2}[\-\/\,\s]{0,3}\w\w\w$")Then
s=rxSelect(s,"^\d+")&"-"&rxSelect(s,"\w+$")&"-"&Year(Date())
s=CDate(s)
ElseIf rxTest(s,"^(Next|Last)(Week|Month|Year)$")Then
s=EXE(s)
ElseIf rxTest(s,"^(Week|Month|Year)(Start|End)$")Then
s=EXE(s&"(Date())")
ElseIf rxTest(s,"\s*[\:\+]\d{4}$")Then
s=rxReplace(s,"\s*[\:\+]\d{4}$","")
s=rxReplace(s,"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\D*","")
s=CDate(s)
ElseIf rxTest(s,"^\d+\-\d+\-\d+(T| )\d+\:\d+")Then
s=rxSelect(s,"^\d+\-\d+\-\d+(T| )\d+\:\d+")
s=Replace(s,"T"," ")
s=CDate(s)
End If
End Select
End If
If Not IsDate(s)Then Exit Function
Dim f,o
f="%D-%b-%y"
s=CDate(s)
o=FmtDate(s,f)
o=UCase(o)
DoDate=o
End Function
Function TimeSince(D1)
TimeSince=TimeGap(D1,Now())
End Function
Function TimeUntil(D2)
TimeUntil=TimeGap(Now(),D2)
End Function
Function TimeGap(D1,D2)
Dim aChunks,aYear,aMonth,aWeek,aDay,aHour,aMinute
Dim dC,aK
Dim iSince
Dim i,j
Dim iSeconds,iSeconds2
Dim sName,sName2
Dim num,num2
Dim sP
Dim Locale
If Not IsDate(D1)Then
TimeGap=""
Exit Function
Else
D1=CDate(D1)
End If
Set dC=Server.CreateObject("Scripting.Dictionary")
dC.Add "year",60 * 60 * 24 * 365
dC.Add "month",60 * 60 * 24 * 30
dC.Add "week",60 * 60 * 24 * 7
dC.Add "day",60 * 60 * 24
dC.Add "hour",60 * 60
dC.Add "minute",60
aK=dC.Keys()
D1=FormatDateTime(D1,0)
If D2="" Then D2=Now()
iSince=CLng(DateDiff("s",D1,D2))
j=dC.Count
For i=0 To j-1
iSeconds=CLng(dC.Item(aK(i)))
sName=aK(i)
num=Int(iSince / iSeconds)
If num<>0 Then
Exit For
End If
Next
If num="-1" Then
sP=sP&"D2"
Else
sP=sP&num&" "&sName&IIf(num>1,"s","")&" ago"
End If
If i+1<j Then
iSeconds2=dC.Item(aK(i+1))
sName2=aK(i+1)
num2=Int((iSince-(iSeconds * num))/ iSeconds2)
If num2<>0 Then
If num2=-1 Then
sP=sP
Else
sP=sP
End If
End If
End If
Set dC=Nothing
Erase aK
TimeGap=sP
End Function
Function DoTime(strDateTime)
On Error Resume Next
strDateTime=strDateTime
If Not(strDateTime<>"")Or(Not IsDate(strDateTime))Then Exit Function
DoTime=FmtDate(strDateTime,"%H:%N")
End Function
Function FormatDate(ByVal strDate,ByVal strFormat)
FormatDate=FmtDate(strDate,strFormat)
End Function
Function FmtDate(ByVal strDate,ByVal strFormat)
strDate=TOE(strDate)
If strDate="" Then Exit Function
strDate=DateFromString(strDate)
If InStr(1,strFormat,"%")>0 Then
Else
strFormat=rxReplace(strFormat,"([a-zA-Z])","%$1")
End If
Dim intPosItem
Dim int12HourPart
Dim str24HourPart
Dim strMinutePart
Dim strSecondPart
Dim strAMPM
strFormat=Replace(strFormat,"%m",DatePart("m",strDate),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%M",Digits(DatePart("m",strDate),2),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%B",MonthName(DatePart("m",strDate),False),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%b",MonthName(DatePart("m",strDate),True),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%d",DatePart("d",strDate),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%D",Digits(DatePart("d",strDate),2),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%O",FmtDayOrdinal(Day(strDate)),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%j",DatePart("y",strDate),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%Y",DatePart("yyyy",strDate),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%y",Right(DatePart("yyyy",strDate),2),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%w",DatePart("w",strDate,1),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%a",WeekDayName(DatePart("w",strDate,1),True),1,-1,vbBinaryCompare)
strFormat=Replace(strFormat,"%A",WeekDayName(DatePart("w",strDate,1),False),1,-1,vbBinaryCompare)
str24HourPart=DatePart("h",strDate)
If Len(str24HourPart)<2 then str24HourPart="0"&str24HourPart
strFormat=Replace(strFormat,"%H",str24HourPart,1,-1,vbBinaryCompare)
int12HourPart=DatePart("h",strDate)Mod 12
If int12HourPart=0 then int12HourPart=12
strFormat=Replace(strFormat,"%h",int12HourPart,1,-1,vbBinaryCompare)
strMinutePart=DatePart("n",strDate)
If Len(strMinutePart)<2 then strMinutePart="0"&strMinutePart
strFormat=Replace(strFormat,"%N",strMinutePart,1,-1,vbBinaryCompare)
If CInt(strMinutePart)=0 then
strFormat=Replace(strFormat,"%n","",1,-1,vbBinaryCompare)
Else
If CInt(strMinutePart)<10 then strMinutePart="0"&strMinutePart
strMinutePart=":"&strMinutePart
strFormat=Replace(strFormat,"%n",strMinutePart,1,-1,vbBinaryCompare)
End if
strSecondPart=DatePart("s",strDate)
If Len(strSecondPart)<2 then strSecondPart="0"&strSecondPart
strFormat=Replace(strFormat,"%S",strSecondPart,1,-1,vbBinaryCompare)
If DatePart("h",strDate)>=12 then
strAMPM="PM"
Else
strAMPM="AM"
End If
strFormat=Replace(strFormat,"%P",strAMPM,1,-1,vbBinaryCompare)
If Err.Number<>0 Then
Err.Clear()
Else
FmtDate=strFormat
End If
End Function
Function FmtDayOrdinal(byVal intDay)
On Error Resume Next
intDay=CInt(intDay)
Dim strOrd
If 1=0 Then
ElseIf intDay=1 Or intDay=21 Or intDay=31 Then
strOrd="st"
ElseIf intDay=2 Or intDay=22 Then
strOrd="nd"
ElseIf intDay=3 Or intDay=23 Then
strOrd="rd"
Else
strOrd="th"
End If
Err.Clear()
FmtDayOrdinal=strOrd
End Function
Function PrintDateAndTime(strDateTime)
PrintDateAndTime=DoDate(strDateTime)&" at "&DoTime(strDateTime)
End Function
Function yymmdd(ByVal dDate)
On Error Resume Next
If Not IsDate(dDate)Then Exit Function
dDate=CDate(dDate)
tDay=Day(dDate)
If Len(tDay)<2 Then tDay="0"&tDay
tMonth=Month(dDate)
If Len(tMonth)<2 Then tMonth="0"&tMonth
tYear=Mid(Year(dDate),3)
yymmdd=tYear&tMonth&tDay
End Function
Const FWX__Date_LikeYears_2010="1802,1813,1819,1830,1841,1847,1858,1869,1875,1886,1897,1909,1915,1926,1937,1943,1954,1965,1971,1982,1993,1999,2010,2021,2027,2038,2049,2055,2066,2077,2083,2094,2100"
Const FWX__Date_LikeYears_2011="1803,1814,1825,1831,1842,1853,1859,1870,1881,1887,1898,1910,1921,1927,1938,1949,1955,1966,1977,1983,1994,2005,2011,2022,2033,2039,2050,2061,2067,2078,2089,2095"
Const FWX__Date_LikeYears_2012="1804,1832,1860,1888,1928,1956,1984,2012,2040,2068,2096"
Const FWX__Date_LikeYears_2013="1805,1811,1822,1833,1839,1850,1861,1867,1878,1889,1895,1901,1907,1918,1929,1935,1946,1957,1963,1974,1985,1991,2002,2013,2019,2030,2041,2047,2058,2069,2075,2086,2097"
Const FWX__Date_LikeYears_2014="1800,1806,1817,1823,1834,1845,1851,1862,1873,1879,1890,1902,1913,1919,1930,1941,1947,1958,1969,1975,1986,1997,2003,2014,2025,2031,2042,2053,2059,2070,2081,2087,2098"
Const FWX__Date_LikeYears_2015="1801,1807,1818,1829,1835,1846,1857,1863,1874,1885,1891,1903,1914,1925,1931,1942,1953,1959,1970,1981,1987,1998,2009,2015,2026,2037,2043,2054,2065,2071,2082,2093,2099"
Const FWX__Date_Dec1st_SUN="1805,1811,1822,1833,1839,1850,1861,1867,1878,1889,1895,1901,1907,1918,1929,1935,1946,1957,1963,1974,1985,1991,2002,2013,2019,2030,2041,2047,2058,2069,2075,2086,2097,1816,1844,1872,1912,1940,1968,1996,2024,2052,2080"
Const FWX__Date_Dec1st_MON="1800,1806,1817,1823,1834,1845,1851,1862,1873,1879,1890,1902,1913,1919,1930,1941,1947,1958,1969,1975,1986,1997,2003,2014,2025,2031,2042,2053,2059,2070,2081,2087,2098,1828,1856,1884,1924,1952,1980,2008,2036,2064,2092"
Const FWX__Date_Dec1st_TUE="1801,1807,1818,1829,1835,1846,1857,1863,1874,1885,1891,1903,1914,1925,1931,1942,1953,1959,1970,1981,1987,1998,2009,2015,2026,2037,2043,2054,2065,2071,2082,2093,2099,1812,1840,1868,1896,1908,1936,1964,1992,2020,2048,2076"
Const FWX__Date_Dec1st_WED="1802,1813,1819,1830,1841,1847,1858,1869,1875,1886,1897,1909,1915,1926,1937,1943,1954,1965,1971,1982,1993,1999,2010,2021,2027,2038,2049,2055,2066,2077,2083,2094,2100,1824,1852,1880,1920,1948,1976,2004,2032,2060,2088"
Const FWX__Date_Dec1st_THU="1803,1814,1825,1831,1842,1853,1859,1870,1881,1887,1898,1910,1921,1927,1938,1949,1955,1966,1977,1983,1994,2005,2011,2022,2033,2039,2050,2061,2067,2078,2089,2095,1808,1836,1864,1892,1904,1932,1960,1988,2016,2044,2072"
Const FWX__Date_Dec1st_FRI="1809,1815,1826,1837,1843,1854,1865,1871,1882,1893,1899,1905,1911,1922,1933,1939,1950,1961,1967,1978,1989,1995,2006,2017,2023,2034,2045,2051,2062,2073,2079,2090,1820,1848,1876,1916,1944,1972,2000,2028,2056,2084"
Const FWX__Date_Dec1st_SAT="1810,1821,1827,1838,1849,1855,1866,1877,1883,1894,1900,1906,1917,1923,1934,1945,1951,1962,1973,1979,1990,2001,2007,2018,2029,2035,2046,2057,2063,2074,2085,2091,1804,1832,1860,1888,1928,1956,1984,2012,2040,2068,2096"
Function DateIsBefore(ByVal a,ByVal b)
If a<>"" And b<>"" And CDate(DoDate(a))<CDate(DoDate(b))Then
DateIsBefore=True
Else
DateIsBefore=False
End If
End Function
Function DateIsAfter(ByVal a,ByVal b)
If a<>"" And b<>"" And CDate(DoDate(a))>CDate(DoDate(b))Then
DateIsAfter=True
Else
DateIsAfter=False
End If
End Function
Function DateIs(ByVal a,ByVal b)
If a<>"" And b<>"" And DoDate(a)=DoDate(b)Then
DateIs=True
Else
DateIs=False
End If
End Function
Function IsToday(d)
IsToday=DateIs(Date(),d)
End Function
Function IsTomorrow(d)
IsTomorrow=DateIs(Date()+1,d)
End Function
Function IsYesterday(d)
IsYesterday=DateIs(Date()-1,d)
End Function
Function ColourCodeNumber(ByVal iNumber,ByVal sText)
iNumber=VOZ(iNumber)
If iNumber>0 Then
sColour="#0a0"
ElseIf iNumber<0 Then
sColour="#d00"
Else
sColour="#999"
End If
ColourCodeNumber="<span style=""color:"&sColour&" !important;"">"&TOT(sText,iNumber)&"</span>"
End Function
Function ColourCur(ByVal iNumber)
ColourCur=ColourCodeNumber(iNumber,Cur(iNumber))
End Function
Function ColourNum(ByVal iNumber)
ColourNum=ColourCodeNumber(iNumber,iNumber)
End Function
Function ColourCodeStatus(ByVal str)
Select Case Trim(LCase(str))
Case "active","yes","true"
output="<span style=""color:green;"">"&str&"</span>"
Case "paused","disabled","disactive"
output="<span style=""color:blue;"">"&str&"</span>"
Case "deleted","stopped","inactive","removed"
output="<span style=""color:red;"">"&str&"</span>"
Case Else
output="<span style="""">"&str&"</span>"
End Select
ColourCodeStatus=output
End Function
Function ColourCodeFlag(ByVal bln,ByVal str)
If IsTrue(bln)Then
output="<span style=""color:green;"">"&str&"</span>"
Else
output="<span style=""color:red;"">"&str&"</span>"
End If
ColourCodeFlag=output
End Function
Function ColourYesNo(ByVal bln)
ColourYesNo=ColourCodeFlag(bln,YesNo(bln))
End Function
Function Digits(ByVal s,ByVal n)
s=s&""
If IsEmpty(s)Then Exit Function
If n>0 And n>Len(s)Then
s=String(n-Len(s),"0")&s
End if
Digits=s
End Function
Function MakeVar(ByVal str,ByVal KeepLineFeeds)
str=Encode(str)
str=Replace(str,Chr(10),"")
If KeepLineFeeds Then Replacement=""" & vbCRLf & """ Else Replacement=Empty
str=Replace(str,vbCRLf,Replacement)
MakeVar=str
End Function
Function YesNo(bln)
If IsTrue(bln)Then
YesNo="yes"
Else
YesNo="no"
End If
End Function
Function PlusMinus(ByVal iNumber)
If GTZ(iNumber)Then
PlusMinus="+"
Else
PlusMinus="-"
End If
End Function
Function Enclose(ByVal sString,ByVal sBefore,ByVal sAfter)
sString=sString
If sString<>"" Then
sString=sBefore&sString&sAfter&""
End If
Enclose=sString
End Function
Function PutBefore(ByVal sString,ByVal sBefore)
PutBefore=Enclose(sString,sBefore,"")
End Function
Function PutAfter(ByVal sString,ByVal sAfter)
PutAfter=Enclose(sString,"",sAfter)
End Function
Function Text2Html(ByVal s)
s=Trim(s)
If s<>"" Then
Dim x
x="TEXT2HTML_ENCODED_CHAR_DELIMITER"
s=rxReplace(s,"&#([^;]+);","&#$1"&x)
Dim a
a="TEXT2HTML_ENCODED_CHAR_AMP"
s=Replace(s,"&amp;",a)
s=rxReplace(s,"\s+-\s+;",";")
s=rxReplace(s,"\;[\r\n]*$","")
s=rxReplace(s,"(\;[\r\n]*)+","$1")
s=rxReplace(s,"\;[\r\n]*","<br/>")
s=Replace(s,vbCRLf,"<br/>")
s=Replace(s,";","<br/>")
s=rxReplace(s,"(br/>)[\s\r\n]+","$1")
s=rxReplace(s,"(<br/>)+","<br/>")
s=rxReplace(s,"\-\s*;",";")
s=rxReplace(s,"\-$","")
s=Replace(s,a,"&amp;")
s=Replace(s,x,";")
End If
Text2Html=s
End Function
Function Encode(ByVal s)
s=HTMLEncode(s)
s=QuotesEncode(s)
Encode=s
End Function
Function Decode(ByVal s)
s=HTMLDecode(s)
s=QuotesDecode(s)
Decode=s
End Function
Function STRING_Encode(ByVal s):STRING_Encode=Encode(s):End Function
Function STRING_Decode(ByVal s):STRING_Decode=Decode(s):End Function
Function StripMarkup(ByVal s):StripMarkup=StripHtml(s):End Function
Function StripHtml(ByVal s)
s=rxReplace(s,"<[^>]*?>","")
StripHtml=s
End Function
Function HTMLEncode(ByVal s)
s=s
If s<>"" Then
s=replace(s,"&","&amp;")
s=replace(s,">","&gt;")
s=replace(s,"<","&lt;")
End If
HTMLEncode=s
End Function
Function HTMLDecode(ByVal s)
s=s
If s<>"" Then
s=Replace(s,"&amp;","&")
s=Replace(s,"&lt;","<")
s=Replace(s,"&gt;",">")
s=Replace(s,"&nbsp;"," ")
Dim I
For I=1 to 255
s=Replace(s,"&#"&I&";",Chr(I))
Next
End If
HTMLDecode=s
End Function
Function QuotesEncode(ByVal s)
s=s
If s<>"" Then
s=Replace(s,"'","&#39;")
s=Replace(s,"""","&quot;")
End If
QuotesEncode=s
End Function
Function QuotesDecode(ByVal s)
s=s&""
If s<>"" Then
s=replace(s,"&#39;","'")
s=replace(s,"&quot;","""")
End If
QuotesDecode=s
End Function
Function EncodeTSV(ByVal s):EncodeTSV=TSVEncode(s):End Function
Function TSVEncode(ByVal s)
s=s&""
s=Replace(s,Chr(9)," ")
s=Replace(s,Chr(10)&Chr(13),"\n")
s=Replace(s,Chr(10),"\n")
s=Replace(s,Chr(13),"\n")
TSVEncode=s
End Function
Function EncodeTXT(ByVal s):EncodeTXT=TXTEncode(s):End Function
Function TXTEncode(ByVal s)
s=s&""
s=Replace(s,vbTab," ")
s=rxReplace(s,"[\r\n]",";")
TXTEncode=s
End Function
Function EncodeJS(ByVal s):EncodeJS=JSEncode(s):End Function
Function JSEncode(ByVal s)
s=s&""
s=Replace(s,"\","\\")
s=Replace(s,"'","\'")
s=Replace(s,"""","\""")
s=rxReplace(s,"\r\n","\n")
s=rxReplace(s,"[\r\n]","\n")
s=rxReplace(s,"\t","\t")
s=rxReplace(s,"\s"," ")
JSEncode=s
End Function
Function EncodeJSON(ByVal s):EncodeJSON=JSONEncode(s):End Function
Function JSONEncode(ByVal s)
s=s&""
s=Replace(s,"\","\\")
s=Replace(s,"""","\""")
s=rxReplace(s,"\r\n","\n")
s=rxReplace(s,"[\r\n]","\n")
s=rxReplace(s,"\t","\t")
s=rxReplace(s,"\s"," ")
JSONEncode=s
End Function
Function EncodeXml(ByVal s):EncodeXml=XmlEncode(s):End Function
Function XmlEncode(ByVal s)
s=Encode(s)
XmlEncode=s
End Function
Function XmlDecode(ByVal s)
s=Decode(s)
XmlDecode=s
End Function
Function XmlEscape(s)
s=XmlEncode(s)
XmlEscape=s
End Function
Function XmlUnescape(s)
s=XmlDecode(s)
XmlUnescape=s
End Function
Function Escape(s)
Dim i,o
For i=1 To Len(s)
o=o&"%"&Right("0"& Hex(Asc(Mid(s,i,1))),2)&""
Next
Escape=o
End Function
Function Unescape(s)
Unescape=URLDecode(s)
End Function
Function UrlEncode(ByVal s)
s=s
If s<>"" Then
s=Server.URLEncode(s)
End If
UrlEncode=s
End Function
Function URLDecode(ByVal s)
s=s
If s<>"" Then
Dim i,h,ce,cd,valid
Set valid=New RegExp
valid.Pattern="[\w\d]{2}"
s=Replace(s,"+"," ")
s=Replace(s,"%22",Chr(34))
i=InStr(1,s,"%")
Do While i>0 And(i+2)<=Len(s)
ce=Mid(s,i+1,2)
h="&H"&ce&""
If valid.Test(ce)Then
Err.Clear()
cd=CStr(Chr(CLng(h)))
If Err.Description<>"" Or Err.Number<>0 Then cd=""
Err.Clear()
End If
If cd<>"" Then
If cd="%" Then cd="-!-"
Else
cd="-!-"&ce
End If
s=Replace(s,"%"&ce,cd)
h="":cd="":ce=""
i=InStr(1,s,"%")
Loop
s=Replace(s,"-!-","%")
End If
URLDecode=s
End Function
Function Dec2Hex(ByVal DecVal)
Dec2Hex=Hex(DecVal)
End Function
function hex2dec(byVal value)
hex2dec=0
dim i,u,d
i=0:u=len(value)
while(u>=1)
d=instr("0123456789ABCDEF",mid(value,u,1))
hex2dec=hex2dec+((d-1)*(16 ^ i))
i=i+1:u=u-1
wend
end function
Function UnicodeToAscii(ByRef pstrUnicode)
Dim llngLength
Dim llngIndex
Dim llngAscii
Dim lstrAscii
llngLength=Len(pstrUnicode)
For llngIndex=1 To llngLength
llngAscii=Asc(Mid(pstrUnicode,llngIndex,1))
lstrAscii=lstrUnicode&ChrB(llngAscii)
Next
UnicodeToAscii=lstrAscii
End Function
Function ASCII(ByVal s)
ASCII=rxReplace(s&"",rxNonASCII,"")
End Function
Function ASCIICharacters()
Dim a,b,c,o,h
h="0123456789ABCDEF"
Set o=New QuickString
For a=1 To 16
For b=1 To 16
c=CStr(Chr("&H"&Mid(h,a,1)&Mid(h,b,1)&""))
If rxTest(c,rxASCII)Then
o.Add c
End If
Next
Next
ASCIICharacters=o
End Function
Const HEX_Delimiter=" "
Function HexEncode(ByVal s)
Dim x
For x=1 To Len(s)
HexEncode=HexEncode&Right(String(2,"0")& Hex(Asc(Mid(s,x,1))),2)
If x<Len(s)Then HexEncode=HexEncode& HEX_Delimiter
Next
End Function
Function HexDecode(ByVal s)
For Each x In Split(s,HEX_Delimiter)
HexDecode=HexDecode&Chr("&H"&x)
Next
End Function
Function FixHTML_UserTemplate(ByVal s)
s=RemoveComments(s)
s=rxReplace(s,"<p[^>]*?>[\n\r\s\t]*"&"<(div|table|script)\b","<$1")
s=rxReplace(s,"</(div|table|script)>"&"[\n\r\s\t]*</p>","</$1>")
s=rxReplace(s,"<p[^>]*?>[\n\r\s\t]*"&"((<span[^>]*?>)[\n\r\s\t]*)*"&"<(div|table)\b","$1<$3")
s=rxReplace(s,"</(div|table)>"&"[\n\r\s\t]*(</span>"&"[\n\r\s\t]*)*</p>","</$1>$2")
s=rxReplace(s,"<p[^>]*?>[\n\r\s\t]*(&nbsp;)?[\n\r\s\t]*</p>","")
FixHTML_UserTemplate=s
End Function
Function FixHtml(ByVal s)
s=s&""
If IsFront Then
s=PermanentRedirectAnticipate(s)
End If
s=Replace(s,"�","")
s=Replace(s,"�","")
s=Replace(s,"�","")
s=Replace(s,"�","")
s=Replace(s,"�","")
s=Replace(s,"�","")
s=Replace(s,"�","")
s=Replace(s,"�","")
s=RemoveComments(s)
r.Pattern="<script((?!<\/script>)[\s\S])*?<\/script>|<style((?!</style>)[\s\S])*?<\/style>"
Dim sDNTouch:sDNTouch="$~DNTouch-$"
Dim aDNTouch:Set aDNTouch=r.Execute(s)
s=r.Replace(s,sDNTouch)
s=HTMLEntities_PutBack(s)
s=FixPaging(s)
s=FixAmps(s)
s=FixPound(s)
s=rxReplace(s,"<(col|img|input|br|hr|meta)\b([^>]*?[^/])?>","<$1$2/>")
s=rxReplace(s,"<(span|div|strong|em|p)([^>]+)/>","<$1$2></$1>")
s=rxReplace(s,"gtbfieldid=""\d+""\s+","")
s=rxReplace(s,"/\.(js|json|doc|xml|pdf|csv|txt|tsv|xls)",".$1")
DBG "NORMALIZE DOMAINS"
DBG "NORMALIZE DOMAINS> MyDomain="&MyDomain
DBG "NORMALIZE DOMAINS> SERVER_NAME="&SERVER_NAME
DBG "NORMALIZE DOMAINS> DOMAIN_NAME="&DOMAIN_NAME
Dim d,TrailingSlash:TrailingSlash="/"
Dim dn:dn=FWX.Setting("fixhtml.normalize.domains")
Dim ds:ds=FWX.Setting("fixhtml.normalize.subdomains")
DBG "NORMALIZE DOMAINS> dn(known domains)="&dn
DBG "NORMALIZE DOMAINS> ds(known sub-dms)="&ds
Dim d0:Set d0=Dictionary()
For Each d In Split(DOMAIN_NAME&","&MyDomain&","&SERVER_NAME&","&dn,",")
If d<>"" Then
If d0(d)<>"" Then
Else
DBG "NORMALIZE> DOMAIN>"&d&" (first run)"
d0(d)=1
s=rxReplace(s,"(href|src|action)\=(""|')https?://((www|local|l\d|temp|test|dev|"&FWX__BO_Reserved_SubDomains&"|m|d\d|w\d|v\d)\.)*"&rxEscape(d)&"[^/""']*/?","$1=$2"&TrailingSlash)
If ds<>"" Then
DBG "NORMALIZE> SUB-DOMAIN>"&ds&"."&d
s=rxReplace(s,"(href|src|action)\=(""|')https?://("&ds&")"&rxEscape(d)&"[^/]*/","$1=$2"&TrailingSlash)
End If
End If
End If
Next
s=rxReplace(s,"(href|src|action)\=(""|')https?://"&"localhost"&"[^/]*/","$1=$2"&TrailingSlash)
s=Replace(s,"-fixed."&Mydomain,"."&Mydomain)
s=rxReplace(s,"(https?)://fixed-(sub(domain)-)?","$1://")
s=rxReplace(s,"(href|src|action)\=(""|')/([^""'\?]+)/([""'\?])","$1=$2/$3$4")
Dim iA,iB,align,iC,iQ
For Each iA In rxMatches(s,"<img((?!align)[^><])*?v?align=[\""\']\w+[\""\'][^><]*?>")
align=rxSelect(iA,"align=[\""\']\w+[\""\']")
align=rxReplace(align,"^align=[\""\']","")
align=rxReplace(align,"[\""\']$","")
If rxTest(align,"top|middle|bottom")Then
align="vertical-align:"&align
Else
align="float:"&align
End If
iC=rxReplace(iA,"v?align=[\""\']\w+[\""\']","")
iC="<span style='"&align&"'>"&iC&"</span>"
iC=iC&"<!--FixHTML-->"
s=Replace(s,iA,iC)
Next
s=rxReplace(s,"\s+</a>(\w)","</a> $1")
s=rxReplace(s,"\s+</a>","</a>")
s=rxReplace(s,"<p[^>]*?>[\n\r\s\t]*"&"<(div|table|script)\b","<$1")
s=rxReplace(s,"</(div|table|script)>"&"[\n\r\s\t]*</p>","</$1>")
s=CrunchHTML(s)
i=0
Do While InStr(1,s,sDNTouch)>0
s=Replace(s,sDNTouch,aDNTouch(i),1,1)
i=i+1
Loop
aDNTouch="":sDNTouch=""
s=rxReplace(s,"\s*(</(body|html)>)\s*","$1")
s=rxReplace(s,"(href|src|action)\=(""|')\/([^\/])","$1=$2"&"//"&SERVER_NAME&"/$3")
s=Replace(s,"http://localhost/","http://"&SERVER_NAME&"/")
FixHtml=s
End Function
Function RemoveComments(ByVal s)
r.Pattern="<!-"&"-\[if((?!endif\]-->)[\s\S])*endif\]-"&"->|<script((?!<\/script>)[\s\S])*?<\/script>|<style((?!</style>)[\s\S])*?<\/style>"
Dim sDNTouch:sDNTouch="$!$"
Dim aDNTouch:Set aDNTouch=r.Execute(s)
s=r.Replace(s,sDNTouch)
Dim ms,mi,rp,ft,cs,ci
cs=0
ft="DelayedRemove"
Set ms=rxMatches(s,"<!-"&"-[^\[]((?!-->)[\s\S])*-"&"->")
For Each mi In ms
rp=""
If InStr(1,mi,"$!$")>0 Then
cs=cs+1
rp="~~~CommentedOutScript"&cs&"~~~>"
j(rp)=mi
End If
s=Replace(s,mi,rp,1,1)
Next
Set ms=Nothing
ci=1
Do While ci<=cs
n="~~~CommentedOutScript"&ci&"~~~>"
s=Replace(s,n,j(n),1,1)
ci=ci+1
Loop
ci=0
Do While InStr(1,s,sDNTouch)>0
s=Replace(s,sDNTouch,aDNTouch(ci),1,1)
ci=ci+1
Loop
aDNTouch="":sDNTouch=""
s=rxReplace(s,"\{"&ft&"}"&"[^\[]((?!"&"/"&ft&"\}"&")[\s\S])*?"&"\{\/"&ft&"\}","")
RemoveComments=s
End Function
Function FixXml(ByVal s)
s=Replace(s,"\x96","-")
s=Replace(s,"\x92","'")
s=Replace(s,"`","'")
s=Replace(s,"’","'")
FixXml=s
End Function
Public Function FixPaging(ByVal s)
If IsFront Then
Dim rxP,a,aP,iPG,iSZ
rxP="(\.(asp|html?)|/)"&"\?\@(size|page)\=([0-9]+)\&(amp;)?\@(size|page)\=([0-9]+)\&(amp;)?"
If rxTest(s,rxP)Then
s=rxReplace(s,"/((show|page)[0-9]+)(-?)(show[0-9]+)?"&"(\.(asp|html?)|/)"&"(?=(\?\@(size|page)))","/")
Set aP=rxMatches(s,rxP)
For Each a In aP
iPG=VOV(rxFind(rxFind(a,"\@page\=[0-9]+"),"[0-9]+"),1)
iSZ=VOV(rxFind(rxFind(a,"\@size\=[0-9]+"),"[0-9]+"),25)
s=Replace(s,a,"/page"&iPG&"-show"&iSZ&"/?")
Set aP=rxMatches(s,rxP)
Next
s=rxReplace(s,"/page1\-?(show[0-9]+)?/","/$1/")
s=rxReplace(s,"/(page[0-9]+)?-?show25/","/$1/")
s=rxReplace(s,"\/\?(""|')","/$1")
End If
End If
FixPaging=s
End Function
Public Function FixAmps(ByVal s)
s=s&""
s=rxReplace(s,"&((?!(\#x?\w{1,4}|\w{2,6});).)","&amp;$1")
FixAmps=s
End Function
Public Function FixPound(ByVal s)
s=s&""
FixPound=s
End Function
Function CrunchHTML(ByVal s)
s=s&""
r.Pattern="<pre((?!<\/pre>)[\s\S])*?<\/pre>|<textarea((?!<\/textarea>)[\s\S])*?<\/textarea>"
Dim sDNCrunch:sDNCrunch="$~DNCrunch-$"
Dim aDNCrunch:Set aDNCrunch=r.Execute(s)
s=r.Replace(s,sDNCrunch)
s=rxReplace(s,"[ \t]+"," ")
s=rxReplace(s,"\s+(\/)*(\/)?\>","$1$2>")
s=rxReplace(s,"[\n\r\s]{2,}",vbCRLf)
s=rxReplace(s,"(\s*[\n\r]+\s*)+",vbCRLf)
s=rxReplace(s,"^[\r\n]+|[\r\n]+$","")
s=Trim(s)
Dim i:i=0
Do While InStr(1,s,sDNCrunch)>0
s=Replace(s,sDNCrunch,aDNCrunch(i),1,1)
i=i+1
Loop
aDNCrunch="":sDNCrunch=""
CrunchHTML=s
End Function
Function HighlightPhrases(ByVal s,ByVal w)
w=UniquePhrases(w)
HighlightPhrases=HighlightSegments(s,w,",",0)
End Function
Function HighlightWords(ByVal s,ByVal w)
w=UniqueWords(w)
HighlightWords=HighlightSegments(s,w," ",0)
End Function
Function Highlight(ByVal s,ByVal w)
w=SearchWords(Trim(DecodeCriteria(w)))
If w="" Then
w=SearchWords(Trim(DecodeCriteria(Query.FQ("q")&" "&Query.Data("_REFQ"))))
End If
Highlight=HighlightSegments(s,w," ",5)
End Function
Function HighlightSegments(ByVal s,ByVal w,ByVal delimiter,ByVal colours)
Const TAG_MARKER="$*%"
Const ENT_MARKER="@#@"
Const HEAD_MRKER="%@$"
Const NOHI_MRKER=":-!"
Const DNH_MARKER="DO_NOT_HIGHLIGHT"
s=s&"":w=w&"":colours=VOZ(colours)
If w<>"" Then
s=rxReplace(s,"(&#?)(\w+);","<"&DNH_MARKER&">$1$2;</"&DNH_MARKER&">")
r.Pattern="<head((?!</head>)[\s\S])*?<\/head>"
If r.Test(s)Then
Set aTags=r.Execute(s)
sHead=aTags(0)
s=r.Replace(s,HEAD_MRKER)
End If
r.Pattern="<script((?!<\/script>)[\s\S])*?<\/script>|<style((?!</style>)[\s\S])*?<\/style>|<title((?!</title>)[\s\S])*?<\/title>|<button((?!</button>)[\s\S])*?<\/button>|<select((?!</select>)[\s\S])*?<\/select>|<option((?!</option>)[\s\S])*?<\/option>|<textarea((?!</textarea>)[\s\S])*?<\/textarea>|<strong((?!</strong>)[\s\S])*?<\/strong>|<embed((?!</embed>)[\s\S])*?<\/embed>|<"&DNH_MARKER&"((?!</"&DNH_MARKER&">)[\s\S])*?<\/"&DNH_MARKER&">|zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
Set aNoHi=r.Execute(s)
s=r.Replace(s,NOHI_MRKER)
r.Pattern="<((?!(>))[\s\S])+?(>)"
Set aTags=r.Execute(s)
s=r.Replace(s,TAG_MARKER)
aWords=Split(w,delimiter)
For i=0 To UBound(aWords)
If Trim(aWords(i)&"")<>"" Then
r.Pattern="("&rxEscape(aWords(i))&")"
cls="H0"
If colours>0 Then cls="H"&((i Mod colours)+1)
s=r.Replace(s,"~@#"&cls&":@:$1#@~")
End If
Next
i=0
Do While InStr(1,s,TAG_MARKER)>0
s=Replace(s,TAG_MARKER,aTags(i),1,1)
i=i+1
Loop
i=0
Do While InStr(1,s,NOHI_MRKER)>0
s=Replace(s,NOHI_MRKER,aNoHi(i),1,1)
i=i+1
Loop
s=Replace(s,HEAD_MRKER,sHead)
s=rxReplace(s,"<"&DNH_MARKER&">(\#|\&)(\w+);</"&DNH_MARKER&">","$1$2;")
s=Replace(s,"<"&DNH_MARKER&">","")
s=Replace(s,"</"&DNH_MARKER&">","")
s=Replace(s,"<DO_NOT_HIGHLIGHT>","")
s=Replace(s,"</DO_NOT_HIGHLIGHT>","")
r.Pattern="~@#"&"(H\d+)"&":@:([^#]+)#@~"
s=r.Replace(s,"<strong class='H"&"$1"&"'>"&"$2</strong>")
End If
HighlightSegments=s
End Function
Function AutoLink(ByVal s)
AutoLink=s:Exit Function
AutoLink=s
s=s&""
r.Pattern="<head((?!</head>)[\s\S])*?<\/head>"
Dim sHEAD:sHEAD="<!--%AL!HEAD!AL%-->"
Dim aHEAD:Set aHEAD=r.Execute(s)
s=r.Replace(s,sHEAD)
r.Pattern="(<script((?!<\/script>)[\s\S])*?<\/script>)|(<style((?!</style>)[\s\S])*?<\/style>)|(<title((?!</title>)[\s\S])*?<\/title>)|(<button((?!</button>)[\s\S])*?<\/button>)|(<select((?!</select>)[\s\S])*?<\/select>)|(<option((?!</option>)[\s\S])*?<\/option>)|(<textarea((?!</textarea>)[\s\S])*?<\/textarea>)|(<a((?!</a>)[\s\S])*?<\/a>)"
Dim sTAGS:sTAGS="<!--%AL!TAGS!AL%-->"
Dim aTAGS:Set aTAGS=r.Execute(s)
s=r.Replace(s,sTAGS)
r.Pattern="(<((?!>)[\s\S])*?>)+"
Dim sHTML:sHTML="<!--%AL!HTML!AL%-->"
Dim aHTML:Set aHTML=r.Execute(s)
s=r.Replace(s,sHTML)
If 1=0 Then
ElseIf EXT<>"" And(Not rxTest(EXT,"html?|asp"))Then
ElseIf IsTrue(FWX.Setting("no-autolink"))Then
ElseIf IsSet(j("no-autolink"))Then
Else
Dim ptn,rpl,url,lbl,cls,rel,tgt,typ,don
don="":ptn="("&rxEmailAddress&")|("&rxURL&")"
Dim pot,hld,lnk,i
Set lnk=Dictionary()
pot=s
url=rxSelect(pot,ptn)
i=0
Do While url<>""
i=i+1
If InStr(1,don,","&url&",")>0 Then
Else
don=don&","&url&","
typ=IIf(rxTest(url,rxEmailAddress),"email","url")
If Not IsTrue(FWX.Setting("no-autolink-"&typ))Then
lbl=url
ref=url
tgt=IIf(InStr(1,url,MyDomain)>0,"_top","_blank")
rel=IIf(InStr(1,url,MyDomain)>0,"internal","nofollow")
cls=IIf(InStr(1,url,MyDomain)>0,"internal","external")
If rxTest(url,rxEmail)Then
ref="mailto:"&url
Else
ref=IIf(rxTest(url,"^("&rxProtocol&")"),"","http://")&url
don=don&","&ref&","
End If
rpl="<a href="""&ref&""" target="""&tgt&""" rel="""&rel&""" class=""auto-link auto-link-"&cls&" auto-link-"&typ&""">"&lbl&"</a>"
hld="AUTO-LINK-("&i&")"
lnk(hld)=rpl
s=Replace(s,url,hld)
End If
End If
pot=Replace(pot,url,"***AUTO-LINKED(A)***")
pot=rxReplace(pot,"(<a((?!</a>)[\s\S])*?<\/a>)","***AUTO-LINKED(B)***")
url=rxSelect(pot,ptn)
Loop
For Each hld In lnk
s=Replace(s,hld,lnk(hld))
Next
End If
i=0:Do While InStr(1,s,sHTML)>0
s=Replace(s,sHTML,aHTML(i),1,1)
i=i+1:Loop:sHTML="":aHTML=""
i=0:Do While InStr(1,s,sTAGS)>0
s=Replace(s,sTAGS,aTAGS(i),1,1)
i=i+1:Loop:sTAGS="":aTAGS=""
i=0:Do While InStr(1,s,sHEAD)>0
s=Replace(s,sHEAD,aHEAD(i),1,1)
i=i+1:Loop:sHEAD="":aHEAD=""
AutoLink=s
End Function
Function PCase(ByVal str)
Dim iPosition,iSpace,output
iPosition=1
Do While InStr(iPosition,str," ",1)<>0
iSpace=InStr(iPosition,str," ",1)
output=output&UCase(Mid(str,iPosition,1))
output=output&LCase(Mid(str,iPosition+1,iSpace-iPosition))
iPosition=iSpace+1
Loop
output=output&UCase(Mid(str,iPosition,1))
output=output&LCase(Mid(str,iPosition+1))
PCase=output
End Function
Function Preview(ByVal str,ByVal q)
MyWords=Split(q," ")
Matches=0
output=Empty
i=0
Do While(i<=UBound(MyWords))
MatchPosition=InStr(LCase(str),Trim(LCase(MyWords(i))))
If MatchPosition>0 Then
If MatchPosition>30 Then
output=output&"..."&Trim(Mid(str,MatchPosition-30,60))
If Len(str)>(MatchPosition+30)Then output=output&"..."
Else
output=output&Trim(Left(str,60))
If Len(str)>Len(output)Then output=output&"..."
End If
output=output&"&nbsp;"
End If
i=i+1
Loop
If output=Empty Then
output=Left(str,60)
If Len(str)>Len(output)Then output=output&"..."
End If
Preview=Highlight(output,q)
End Function
Function Reverse(ByVal str)
If isEmpty(str)Then Exit Function
i=1:output=Empty
Do While i<=Len(str)
output=Mid(str,i,1)&output
i=i+1
Loop
Reverse=output
End Function
Function StartsWith(ByVal s,ByVal q)
StartsWith=CBool(Left(s,Len(q))=q)
End Function
Function EndsWith(ByVal s,ByVal q)
EndsWith=CBool(Right(s,Len(q))=q)
End Function
Function Slice(ByVal str,ByVal length)
str=TOE(str)
length=VOZ(length)
If Not Len(str)>0 Then Exit Function
If Not length>0 Then Exit Function
If Len(str)>length+3 Then
tmp=Trim(Left(str,length))&"..."
Else
tmp=Left(str,length+3)
End If
Slice=tmp
End Function
Function BackSlice(ByVal str,ByVal length)
str=TOE(str)
length=VOZ(length)
If Not Len(str)>0 Then Exit Function
If Not length>0 Then Exit Function
If Len(str)>length+3 Then
tmp="..."&Trim(Right(str,length))
Else
tmp=Right(str,length+3)
End If
BackSlice=tmp
End Function
Function UniqueSplit(ByVal s,ByVal d)
Dim p,o
Set o=New QuickString
o.Divider=d
For Each p In Split(s,d)
If p<>"" Then
If Not o.Has(p)Then
o.Add p
End If
End If
Next
UniqueSplit=o
End Function
Function UniquePhrases(ByVal s)
s=rxReplace(s,"(\s*[\r\n\,\;]\s*)+",",")
UniquePhrases=UniqueSplit(s,",")
End Function
Function UniqueWords(ByVal s)
s=rxReplace(s,"(\s*[^\w\-\_]\s*)+"," ")
UniqueWords=UniqueSplit(s," ")
End Function
Function SearchWords(ByVal x)
Dim rq,rqi,ign,s
If x<>"" Then
rq=Split(x," ")
ign="|"&rxBadWords&"|"'" & rxStopWords & "|"
For rqi=0 To UBound(rq)
If InStr(1,ign,"|"&rq(rqi)&"|")>0 Then
rq(rqi)=""
End If
Next
s=Trim(Join(rq," ")&"")
s=Trim(Replace(s,"site:[\w\.\-]+",""))
s=NoDuplicateWords(s)
End If
SearchWords=s
End Function
Function NoDuplicateWords(ByVal x)
If x<>"" Then
Dim s,wrds,wrdi
Set s=New QuickString
s.Divider=" "
wrds=Split(x," ")
For wrdi=0 To UBound(wrds)
If Not s.Has(wrds(wrdi))And IsSet(wrds(wrdi))Then
s.Add wrds(wrdi)
End If
Next
End If
NoDuplicateWords=s
End Function
Function CountWords(x)
CountWords=WordCount(x)
End Function
Function WordCount(x)
Const cstrSingleSpace=" "
Const cstrDoubleSpace="  "
Const cstrCharsToConvert=";,.-"
Dim strX
strX=x
For i=1 To Len(cstrCharsToConvert)
strX=Replace(strX,Mid(cstrCharsToConvert,i,1),cstrSingleSpace)
Next
Do While Instr(strX,cstrDoubleSpace)>0
strX=Replace(strX,cstrDoubleSpace,cstrSingleSpace)
Loop
WordCount=(UBound(Split(strX,cstrSingleSpace)))+1
End Function
Function SliceWords(ByVal str,ByVal num)
SliceWords=SelectWords(str,num)
End Function
Function SelectWords(ByVal str,ByVal num)
Dim prh:prh=str
Dim wrd:wrd="([\w\-]+)"
Dim del:del="(([\!\?\:\=\;\.\s\,\(\)\[\]\/\\\|]+)|(\-\s+))"
If num>0 Then
prh=rxSelect(prh,"^("&wrd&del&"){0,"&(num-1)&"}"&wrd&"*"&del&"?")
End If
SelectWords=prh
End Function
Function SlicePhrases(ByVal str,ByVal num)
SlicePhrases=SelectPhrases(str,num)
End Function
Function SelectPhrases(ByVal str,ByVal num)
Dim prh:prh=str
Dim wrd:wrd="([\w\-]+(([\s\,\(\)\[\]\/\\\|]+)|(\-\s+))?)"
Dim del:del="([\!\?\:\=\;\.]+\s*)"
If num>0 Then
prh=rxSelect(prh,"^("&wrd&"+"&del&"){0,"&(num-1)&"}"&wrd&"*"&del&"?")
End If
SelectPhrases=prh
End Function
Function RandomParagraph(ByVal s)
Dim o,d,a,i
If s<>"" Then
d="-DIVIDER-"
s=rxReplace(s,"[\n\r\s]*<p[^>]*>[\n\r\s]*","")
s=rxReplace(s,"[\n\r\s]*</p[^>]*>[\n\r\s]*",d)
a=Split(s,d)
Randomize()
i=Round(UBound(a)*Rnd())-1
i=IIf(GTZ(i),i,0)
o=a(i)
End If
RandomParagraph=o
End Function
Function DoProfile(ByRef rec,ByVal format)
DoProfile=DoAccount(rec,format)
End Function
Function DoAccount(ByRef rec,ByVal format)
Err.Clear():On Error Resume Next
Dim output:Set output=New QuickString
format=TOT(format,"text")
If Not IsObject(rec)Then Exit Function
Select Case format
Case "plain","0"
output.Add ColItem(rec,"name")
If VOZ(ColItem(rec,"account"))>0 Then output.Add "("&ColItem(rec,"account")&")"
output.Add vbCRLf
If Not IsEmpty(ColItem(rec,"company"))Then output.Add ColItem(rec,"company")&vbCRLf
output.Add vbCRLf&DoAddress(rec,format)
If Not IsEmpty(ColItem(rec,"primary"))Then output.Add DoPhone(ColItem(rec,"primary"))&", "
If Not IsEmpty(ColItem(rec,"alternative"))Then output.Add DoPhone(ColItem(rec,"alternative"))&", "
Case "text","1"
output.Add ColItem(rec,"name")&""
If GTZ(ColItem(rec,"account"))Then output.Add " ("&ColItem(rec,"account")&")"
output.Add ";"&vbCRLf&DoAddress(rec,format)&""
If Not IsEmpty(ColItem(rec,"primary"))Then output.Add ";"&vbCRLf&DoPhone(ColItem(rec,"primary"))&""
If Not IsEmpty(ColItem(rec,"alternative"))Then output.Add ";"&vbCRLf&DoPhone(ColItem(rec,"alternative"))&""
If Not IsEmpty(ColItem(rec,"fax"))Then output.Add ";"&vbCRLf&DoPhone(ColItem(rec,"fax"))&""
If Not IsEmpty(ColItem(rec,"email"))Then output.Add ";"&vbCRLf&ColItem(rec,"email")&""
Case Else
output.Add ColItem(rec,"name")&""
If GTZ(ColItem(rec,"account"))Then output.Add " ("&ColItem(rec,"account")&")"
output.Add "<br/>"&vbCRLf&DoAddress(rec,format)&""
If Not IsEmpty(ColItem(rec,"primary"))Then output.Add "<br/>"&vbCRLf&DoPhone(ColItem(rec,"primary"))&""
If Not IsEmpty(ColItem(rec,"alternative"))Then output.Add "<br/>"&vbCRLf&DoPhone(ColItem(rec,"alternative"))&""
If Not IsEmpty(ColItem(rec,"fax"))Then output.Add "<br/>"&vbCRLf&DoPhone(ColItem(rec,"fax"))&""
If Not IsEmpty(ColItem(rec,"email"))Then output.Add "<br/>"&vbCRLf&ColItem(rec,"email")&""
End Select
DoAccount=output
End Function
Function DoAddress(ByRef rec,ByVal format)
Dim output:Set output=New QuickString
format=TOT(format,"text")
If Not IsObject(rec)Then Exit Function
Dim C_LINE:C_LINE="<LINE>"
Dim C_SECT:C_SECT="<SECT>"
If ColItem(rec,"address1")<>"" Then
If ColItem(rec,"company")<>"" Then output.Add ColItem(rec,"company")&C_LINE
output.Add ColItem(rec,"address1")
If ColItem(rec,"address2")<>"" Then output.Add C_LINE&ColItem(rec,"address2")
If ColItem(rec,"address3")<>"" Then output.Add C_LINE&ColItem(rec,"address3")
If ColItem(rec,"county")<>"" Then output.Add C_LINE&ColItem(rec,"county")
If ColItem(rec,"city")&ColItem(rec,"postcode")<>"" Then
output.Add C_LINE
If ColItem(rec,"city")<>"" Then
output.Add ColItem(rec,"city")
End If
If ColItem(rec,"postcode")<>"" Then
If ColItem(rec,"city")<>"" Then output.Add C_SECT
output.Add UCase(ColItem(rec,"postcode"))
End If
End If
If ColItem(rec,"country")<>"" Then output.Add C_LINE&ColItem(rec,"country")
End If
If output<>"" Then
output=rxReplace(output,C_LINE&"+$","")
output=rxReplace(output,"^"&C_LINE&"+","")
output=rxReplace(output,C_LINE&"+",C_LINE)
output=rxReplace(output,C_SECT&"+",C_SECT)
End If
If output<>"" Then
Select Case format
Case "plain","0"
output=Replace(output,C_LINE,"; ")
output=Replace(output,C_SECT," - ")
Case "text","1"
output=Replace(output,C_LINE,";"&vbCRLf&"")
output=Replace(output,C_SECT," - ")
Case Else
output=Replace(output,C_LINE,"<br/>")
output=Replace(output,C_SECT," - ")
End Select
output=Trim(output)
End If
If output<>"" Then
If Trim(Replace(Replace(output,C_LINE,""),C_SECT,""))="" Then
output=""
End If
End If
DoAddress=output
End Function
Function DoAddress_LEGACY(ByRef rec,ByVal format)
Err.Clear():On Error Resume Next
Dim output:Set output=New QuickString
format=TOT(format,"text")
If Not IsObject(rec)Then Exit Function
Select Case format
Case "plain","0"
If ColItem(rec,"company")<>"" Then output.Add ColItem(rec,"company")&";"&vbCRLf
output.Add ColItem(rec,"address1")
If ColItem(rec,"address2")<>"" Then output.Add ";"&vbCRLf&ColItem(rec,"address2")
If ColItem(rec,"address3")<>"" Then output.Add ";"&vbCRLf&ColItem(rec,"address3")
If ColItem(rec,"county")<>"" Then output.Add ";"&vbCRLf&ColItem(rec,"county")
output.Add ";"&vbCRLf&ColItem(rec,"city")&" - "&UCase(ColItem(rec,"postcode"))
output.Add ";"&vbCRLf&ColItem(rec,"country")
Case "text","1"
If ColItem(rec,"company")<>"" Then output.Add ColItem(rec,"company")&";"&vbCRLf
output.Add ColItem(rec,"address1")
If ColItem(rec,"address2")<>"" Then output.Add ";"&vbCRLf&ColItem(rec,"address2")
If ColItem(rec,"address3")<>"" Then output.Add ";"&vbCRLf&ColItem(rec,"address3")
If ColItem(rec,"county")<>"" Then output.Add ";"&vbCRLf&ColItem(rec,"county")
output.Add ";"&vbCRLf&ColItem(rec,"city")&" - "&UCase(ColItem(rec,"postcode"))
output.Add ";"&vbCRLf&ColItem(rec,"country")
Case Else
If ColItem(rec,"company")<>"" Then output.Add ColItem(rec,"company")&"<br/>"
output.Add ColItem(rec,"address1")
If ColItem(rec,"address2")<>"" Then output.Add "<br/>"&ColItem(rec,"address2")
If ColItem(rec,"address3")<>"" Then output.Add "<br/>"&ColItem(rec,"address3")
If ColItem(rec,"county")<>"" Then output.Add "<br/>"&ColItem(rec,"county")
output.Add "<br/>"&ColItem(rec,"city")&" - "&UCase(ColItem(rec,"postcode"))
output.Add "<br/>"&ColItem(rec,"country")
End Select
DoAddress=output
End Function
Function DoBreakdown(ByRef rsItems,ByRef rsJobs,ByVal format)
Dim output:Set output=New QuickString
format=TOT(format,"text")
output.Add DoItems(rsItems,format)
output.Add DoJobs(rsJobs,format)
DoBreakdown=output
End Function
Function DoPrintInvoice(ByRef rsInv,ByRef p,ByRef s,ByVal format)
Dim output:Set output=New QuickString
format=TOT(format,"text")
Dim sBreakdown
sBreakdown=DoItems(p,format)&DoJobs(s,format)
If sBreakdown="" Then Exit Function
Select Case format
Case "text","0","plain","1"
output.Add sBreakdown
output.Add vbCRLf
If rsInv("discount")>0 Then
output.Add "discount: "&Cur(rsInv("discount"))&";"&vbCRLf
End If
output.Add "total: "&Cur(rsInv("amount"))&";"&vbCRLf
Case "bullet"
Case Else
output.Add "<table width=""100%"">"
If rsInv("discount")>0 Then
output.Add " <tr>"&vbCRLf
output.Add "  <td width=* colspan=2 valign=top align=right>discount:</b></td>"&vbCRLf
output.Add "  <td width=50 valign=top align=right><b>"&Cur(rsInv("discount"))&"</b></td>"&vbCRLf
output.Add " </tr>"&vbCRLf
End If
output.Add " <tr>"&vbCRLf
output.Add "  <td width=* colspan=2 valign=top align=right>total amount:</b></td>"&vbCRLf
output.Add "  <td width=50 valign=top align=right><b>"&Cur(rsInv("amount"))&"</b></td>"&vbCRLf
output.Add " </tr>"&vbCRLf
output.Add "</table>"
End Select
DoPrintInvoice=output
End Function
Function DoInvoice(ByRef rsInv,ByRef p,ByRef s,ByVal format)
DoInvoice=DoPrintInvoice(rsInv,p,s,format)
End Function
Function DoItems(ByRef s,ByVal format)
Dim output:Set output=New QuickString
format=TOT(format,"text")
Dim t:t=""
If Not Exists(s)Then
Exit Function
Else
s.MoveFirst
Select Case format
Case "text","0","plain","1"
Do While Exists(s)
If s("tag")<>t Then
t=s("tag")
output.Add s("tag")&":"&vbCRLf
End If
output.Add s("quantity")&"x "&PCase(s("name"))&" for "&Cur(s("amount"))&";"&vbCRLf
s.MoveNext
Loop
Case "bullet"
Do While Exists(s)
If s("tag")<>t Then
t=s("tag")
output.Add "<div style=""padding:3px 0px 1px 0px;""><u>"&s("tag")&"</u>:</div>"
End If
If IsSet(s("notes"))Then
output.Add "<div style=""padding:3px; border:#F00 solid 1px;"">"
Else
output.Add "<div style=""padding:3px;"">"
End If
output.Add "<b>"&s("quantity")&"</b>x <b>"&PCase(s("name"))&"</b> @ <b>"&Cur(s("amount"))&"</b>"
If IsSet(s("notes"))Then
output.Add "<div style=""padding:4px;"">"&s("notes")&"</div>"
End If
output.Add "</div>"
s.MoveNext
Loop
Case Else
output.Add "<table width=""100%"">"
Do While Exists(s)
If s("tag")<>t Then
t=s("tag")
output.Add "</table>"
output.Add "<div style=""padding:3px 0px 1px 0px;""><u>"&s("tag")&"</u>:</div>"
output.Add "<table width=""100%"">"
End If
output.Add " <tr>"&vbCRLf
output.Add "  <td width=* colspan=2 valign=top><b>"&s("quantity")&"</b>x <b>"&PCase(s("name"))&"</b></td>"&vbCRLf
output.Add "  <td width=50 valign=top align=right><b>"&Cur(s("amount"))&"</b></td>"&vbCRLf
output.Add " </tr>"&vbCRLf
output.Add " <tr>"&vbCRLf
output.Add "  <td colspan=3 style=""padding-left:10px;"" colspan=2>"&vbCRLf
If s("notes")<>Empty Then
output.Add " "&vsp(2)&vbCRLf
output.Add "<div style=""color:blue;"">"&s("notes")&"</div>"
End If
output.Add "  </td>"&vbCRLf
output.Add " </tr>"
output.Add " <tr><td colspan=2>"&vsp(3)&"</td></tr>"&vbCRLf
s.MoveNext
Loop
output.Add "</table>"
End Select
Set s=Nothing
End If
DoItems=output
End Function
Function DoJobs(ByRef s,ByVal format)
Dim output:Set output=New QuickString
format=TOT(format,"text")
If Not Exists(s)Then
Exit Function
Else
s.MoveFirst
Select Case format
Case "text","0","plain","1"
Do While Exists(s)
output.Add PCase(s("name"))&" for "&Cur(s("amount"))&" - "
If IsDate(s("date"))Then
output.Add WeekDayName(WeekDay(s("date")))&" "&DoDate(s("date"))&IIf(IsSet(s("time")),", ","")&s("time")&""
If s("exp_offset")>0 Then
output.Add ", expires "&DoDate(s("date")+s("exp_offset"))&""
End If
End If
output.Add ";"&vbCRLf
s.MoveNext
Loop
Case "bullet"
Do While Exists(s)
If IsSet(s("notes"))Then
output.Add "<div style=""padding:3px; border:#F00 solid 1px;"">"
Else
output.Add "<div style=""padding:3px;"">"
End If
output.Add "<b>"&PCase(s("name"))&"</b> @ <b>"&Cur(s("amount"))&"</b>"
output.Add "<br/>"&vbCRLf
If IsDate(s("date"))Then
output.Add WeekDayName(WeekDay(s("date")))&" "&DoDate(s("date"))&IIf(IsSet(s("time")),", ","")&s("time")&""
If s("exp_offset")>0 Then
output.Add "<br/>"&vbCRLf
output.Add "<div stye=""color:red;"">expires "&DoDate(s("date")+s("exp_offset"))&"</div>"
End If
End If
If IsSet(s("notes"))Then
output.Add "<div style=""padding:4px;"">"&s("notes")&"</div>"
End If
output.Add "</div>"
s.MoveNext
Loop
Case Else
output.Add "<table width=""100%"">"
Do While Exists(s)
output.Add " <tr>"&vbCRLf
output.Add "  <td width=* colspan=2 valign=top><b>"&PCase(s("name"))&"</b></td>"&vbCRLf
output.Add "  <td width=50 valign=top align=right><b>"&Cur(s("amount"))&"</b></td>"&vbCRLf
output.Add " </tr>"&vbCRLf
output.Add " <tr>"&vbCRLf
output.Add "  <td colspan=2 style=""padding-left:10px;"" colspan=2>"&vbCRLf
If IsDate(s("date"))Then
output.Add "   <div class=""Small"">"&WeekDayName(WeekDay(s("date")))&" "&DoDate(s("date"))&IIf(IsSet(s("time")),", ","")&s("time")&"</div>"
If s("exp_offset")>0 Then
output.Add "<div class=""Small"" style=""color:red;"">expires "&DoDate(s("date")+s("exp_offset"))&"</div>"
End If
End If
If s("notes")<>"" Then
output.Add "<div class=""Small"" style=""color:blue;"">"&s("notes")&"</div>"
End If
If s("sg")Then output.Add " &bull; <b class=""Yes"">SG</b>"
If s("fn")Then output.Add " &bull; <b class=""Yes"">FN</b>"
If s("call")Then output.Add " &bull; <b class=""Yes"">CALL</b>"
If s("stairs")Then output.Add " &bull; <b class=""Yes"">STAIRS</b>"
output.Add "  </td>"&vbCRLf
output.Add " </tr>"
output.Add " <tr><td colspan=2>"&vsp(3)&"</td></tr>"&vbCRLf
s.MoveNext
Loop
output.Add "</table>"
End Select
Set s=Nothing
End If
DoJobs=output
End Function
Function DoModifiers(ByRef s,ByVal format)
Dim output:Set output=New QuickString
format=TOT(format,"text")
If Not Exists(s)Then
Exit Function
Else
s.MoveFirst
Select Case format
Case "text","0","plain","1"
Do While Exists(s)
output.Add PCase(s("name"))&" for "&Cur(s("amount"))&";"&vbCRLf
s.MoveNext
Loop
Case "bullet"
Do While Exists(s)
If IsSet(s("description"))Then
output.Add "<div style=""padding:3px; border:#F00 solid 1px;"">"
Else
output.Add "<div style=""padding:3px;"">"
End If
output.Add "<b>"&PCase(s("name"))&"</b> @ <b>"&Cur(s("amount"))&"</b>"
If IsSet(s("description"))Then
output.Add "<div style=""padding:4px;"">"&s("description")&"</div>"
End If
output.Add "</div>"
s.MoveNext
Loop
Case Else
output.Add "<table width=""100%"">"
Do While Exists(s)
output.Add " <tr>"&vbCRLf
output.Add "  <td width=* colspan=2 valign=top><b>"&PCase(s("name"))&"</b></td>"&vbCRLf
output.Add "  <td width=50 valign=top align=right><b>"&Cur(s("amount"))&"</b></td>"&vbCRLf
output.Add " </tr>"&vbCRLf
output.Add " <tr>"&vbCRLf
output.Add "  <td colspan=3 style=""padding-left:10px;"" colspan=2>"&vbCRLf
If s("description")<>Empty Then
output.Add " "&vsp(2)&vbCRLf
output.Add "<H6 style=""color:blue;"">"&s("description")&"</H6>"
End If
output.Add "  </td>"&vbCRLf
output.Add " </tr>"
output.Add " <tr><td colspan=2>"&vsp(3)&"</td></tr>"&vbCRLf
s.MoveNext
Loop
output.Add "</table>"
End Select
Set s=Nothing
End If
DoModifiers=output
End Function
Dim TMPL_PROP:TMPL_PROP="<!--FWX.CACHE-->"&vbCRLf
Const TMPL_PROP_EDITION=0
Const TMPL_PROP_REVISION=1
Const TMPL_PROP_LOCATION=2
Const TMPL_PROP_TEMPLATE=3
Function TMPL__Test(ByVal s):TMPL__Test=CBool(InStr(1,s,TMPL_PROP)):End Function
Function TMPL__Edition(ByVal s):TMPL__Edition=TMPL__Prop(s,TMPL_PROP_EDITION):End Function
Function TMPL__Revision(ByVal s):TMPL__Revision=TMPL__Prop(s,TMPL_PROP_REVISION):End Function
Function TMPL__Location(ByVal s):TMPL__Location=TMPL__Prop(s,TMPL_PROP_LOCATION):End Function
Function TMPL__Template(ByVal s):TMPL__Template=TMPL__Prop(s,TMPL_PROP_TEMPLATE):End Function
Function TMPL__Prop(ByVal s,ByVal i)
Dim x,o
x=Split(s,TMPL_PROP)
If UBound(x)>0 Then
o=x(i)
Else
If i=3 Then s=o
End If
TMPL__Prop=o
End Function
Function TMPL__Empty(ByVal s)
TMPL__Empty=True
If s<>"" Then
If TMPL__Test(s)Then
s=TMPL__Template(s)
If s<>"" Then
TMPL__Empty=False
If StartsWith(s,FWX_CACHE_Empty)Then TMPL__Empty=True
If StartsWith(s,FWX_CACHE_Missing)Then TMPL__Empty=True
End If
Else
TMPL__Empty=False
End If
End If
End Function
Const FWX_CACHE_Empty="<!--EMPTY-->"
Const FWX_CACHE_Missing="<!--MISSING-->"
Const FWX_CACHE_Disabled="<!--DISABLED-->"
Function FWX_Cache_FromRAM(ByVal source)
If source<>"" Then FWX_Cache_FromRAM=FWX_Cache_Valid(Application.Contents(source))("content")
End Function
Function FWX_Cache_FromSession(ByVal source)
If source<>"" Then FWX_Cache_FromSession=FWX_Cache_Valid(Session.Contents(source))("content")
End Function
Function FWX_Cache_FromREDIS(ByVal source)
If source<>"" Then FWX_Cache_FromREDIS=FWX_Cache_Valid(REDIS(source))("content")
End Function
Function FWX_Cache_FromFILE(ByVal source)
If source<>"" Then FWX_Cache_FromFILE=FWX_Cache_Valid(ReadFileWithoutValidation(source))("content")
End Function
Function FWX_Cache_Payload_FromRAM(ByVal source)
If source<>"" Then FWX_Cache_Payload_FromRAM=FWX_Cache_Valid_Payload(Application.Contents(source))
End Function
Function FWX_Cache_Payload_FromSession(ByVal source)
If source<>"" Then FWX_Cache_Payload_FromSession=FWX_Cache_Valid_Payload(Session.Contents(source))
End Function
Function FWX_Cache_Payload_FromREDIS(ByVal source)
If source<>"" Then FWX_Cache_Payload_FromREDIS=FWX_Cache_Valid_Payload(REDIS(source))
End Function
Function FWX_Cache_Payload_FromFILE(ByVal source)
If source<>"" Then FWX_Cache_Payload_FromFILE=FWX_Cache_Valid_Payload(ReadFileWithoutValidation(source))
End Function
Function FWX_Cache_Payload_Make(ByVal content,ByVal control)
FWX_Cache_Payload_Make="<!--{""date"":"""&Now()&""","&IIf(IsDate(control),"""expires"":"""&control&""",",IIf(GTZ(control),"""expires"":"""&VOZ(control)&""",",IIf(Left(control,1)="/","""source"":"""&control&""",",Enclose(control,"""control"":""",""","))))&"""EDITION"":"""&TOT(j("EDITION"),"")&""",""REVISION"":"""&TOT(j("REVISION"),"")&"""}-->"&vbCRLf&_
content&IIf(Right(control,5)=".html",vbCRLf&"<!--END OF "&control&"-->","")&""
End Function
Function FWX_Cache_Payload_Parse(ByVal content)
Dim d
Set d=Dictionary()
If StartsWith(content,FWX_CACHE_Missing)Or StartsWith(content,FWX_CACHE_Disabled)Then
content=""
End If
If content<>"" Then
d("payload")=content
d("meta")=rxSelect(content,"^<!--[^\n\r]+")
d("data")=rxReplace(rxReplace(content,"^<!--[^\n\r]+[\n\r]*",""),"[\n\r]*<!--END OF .*-->$","")
d("data")=rxReplace(d("data"),"^[\t\n\r\s]+|[\t\n\r\s]+$","")
If StartsWith(d("data"),FWX_CACHE_Missing)Or StartsWith(d("data"),FWX_CACHE_Disabled)Then
d("data")=""
End If
If d("meta")<>"" Then
Dim prop,par,val
For Each prop In rxMatches(d("meta"),"""\w+"":""[^""]*""")
par=rxReplace(prop,"^""([^""]+).*","$1")
val=rxReplace(prop,"^""([^""]+)"":""([^""]*).*","$2")
d(par)=val
Next
End If
d("date")=CDate(d("date"))
d("age")=DateDiff("d",d("date"),Now())
End If
Set FWX_Cache_Payload_Parse=d
End Function
Function FWX_Cache_Valid(ByVal payload)
Dim d,o,reason
o=""
If IsObject(payload)Then
Set d=payload
Else
Set d=FWX_Cache_Payload_Parse(payload)
End If
If d("VALIDATED")<>"" Then
ElseIf d("data")="" Then
d("content")=""
Else
d("VALIDATED")="y"
o=d("data")
If d("EDITION")<>j("EDITION")Then
o="":reason="mismatched edition"
End If
If d("REVISION")<>j("REVISION")Then
o="":reason="mismatched revision"
End If
If d("expires")<>"" Then
If IsDate(d("expires"))Then
If CDate(d("expires"))<Now()Then
o="":reason="expired"
Else
End If
ElseIf GTZ(d("expires"))Then
If d("age")>d("expires")Then
o="":reason="expired"
Else
End If
End If
End If
If IsTest Or IsAdmin Then
If d("source")<>"" Then'And Left(d("source"),1)="/" Then 'd("source")<>"none" Then
If reason<>"" Then
Else
d("source-modified")=LastModified(d("source"))
If IsDate(d("source-modified"))Then
If d("source-modified")>d("date")Then
o="":reason="file has been modified"
d("data")=o
d("payload")=""
Else
End If
Else
End If
End If
Else
End If
Else
End If
If IsTest Then
If o="" Then
DBG.Notify "Cache expired "&TOT(d("source"),"/virtual-template/")&": "&reason
Else
If IsObject(MyCookies)Then
If False And IsTest And MyCookies.Global("NoCache.Flag")="y" Then
If d("source")<>"" Then
reason="cache is off and we're testing"
o=ReadFileWithoutValidation(d("source"))
Else
reason="ignoring cache but we don't know the source so can't fetch it on the fly"
o=""
End If
Else
End If
Else
End If
End If
End If
d("ACCEPTED")=CBool(o<>"")
d("content")=o
If o="" Then
d("expired")=True
d("reason")=reason
d("payload")=""
End If
End If
If IsTest Then
End If
Set FWX_Cache_Valid=d
End Function
Function FWX_Cache_Valid_Content(ByVal payload)
FWX_Cache_Valid_Content=FWX_Cache_Valid(payload)("content")
End Function
Function FWX_Cache_Valid_Payload(ByVal payload)
FWX_Cache_Valid_Payload=FWX_Cache_Valid(payload)("payload")
End Function
Class AlphaNumericNumberParser
Private AN__CHARS
Private AN__SCALE
Public Property Get Characters()
Characters=AN__CHARS
End Property
Public Property Let Characters(ByVal s)
AN__CHARS=s
AN__SCALE=Len(s)
End Property
Private Sub Class_Initialize()
Me.Characters="bcdfghjklmnpqrstvwxyz0123456789BCDFGHJKLMNPQRSTVWXYZ"
End Sub
Public Function ToString(ByVal n)
Dim s:s=""
Dim x:x=""
Dim y:y=""
Dim b:b=0
Dim d:d="+"
Do While n>=AN__SCALE
If x<>"" Then x=d&x
b=n Mod AN__SCALE
x=b&x
n=Round((n-b)/AN__SCALE)
Loop
If x<>"" Then x=d&x
x=n&x
For Each y In Split(x,d)
s=s&Mid(AN__CHARS,CInt(y)+1,1)
Next
ToString=s
End Function
Public Function ToNumber(ByVal s)
Dim n:n=0
Dim m:m=Len(s)
Dim k:k=m
Dim b:b=0
Dim u:u=0
Dim i:i=0
Dim c:c=""
Do While m>0
i=k-m
m=m-1
c=Mid(s,i+1,1)
u=Instr(1,AN__CHARS,c)-1
b=(((AN__SCALE-0)^ m)* u)
n=n+b
Loop
ToNumber=n
End Function
End Class
Function HTMLEntities_Remove(ByVal s)
s=rxReplace(s,"&#([\w\d]+);","%~$1%")
HTMLEntities_Remove=s
End Function
Function HTMLEntities_PutBack(ByVal s)
s=rxReplace(s,"%~([\w\d]+)%","&#$1;")
HTMLEntities_PutBack=s
End Function
Function HTMLEntities_Decode(ByVal s)
s=HTMLEntities_PutBack(s)
s=rxReplace(s,"&\#(9|x9);","	")
s=rxReplace(s,"&\#(10|xa);","	")
s=rxReplace(s,"&\#(13|xd);","	")
s=rxReplace(s,"&\#(32|x20);","	")
s=rxReplace(s,"&\#(33|x21);","!")
s=rxReplace(s,"&\#(34|x22);","""")' "	&quot;	" quotation mark
s=rxReplace(s,"&\#(35|x23);","#")
s=rxReplace(s,"&\#(36|x24);","$")
s=rxReplace(s,"&\#(37|x25);","%")
s=rxReplace(s,"&\#(38|x26);","&")
s=rxReplace(s,"&\#(39|x27);","'")
s=rxReplace(s,"&\#(40|x28);","(")
s=rxReplace(s,"&\#(41|x29);",")")
s=rxReplace(s,"&\#(42|x2a);","*")
s=rxReplace(s,"&\#(43|x2b);","+")
s=rxReplace(s,"&\#(44|x2c);",",")
s=rxReplace(s,"&\#(45|x2d);","-")
s=rxReplace(s,"&\#(46|x2e);",".")
s=rxReplace(s,"&\#(47|x2f);","/")
s=rxReplace(s,"&\#(163|xa3);","&pound;")
s=Replace(s,"‘","&lsquo;")
s=Replace(s,"’","&rsquo;")
s=Replace(s,"“","&ldquo;")
s=Replace(s,"”","&rdquo;")
s=Replace(s,"‚","&sbquo;")
s=Replace(s,"„","&bdquo;")
s=Replace(s,"′","&prime;")
s=Replace(s,"″","&Prime;")
s=Replace(s,"‐","&#8208;")
s=Replace(s,"—","&#8208;")
s=Replace(s,"–","&ndash;")
s=Replace(s,"—","&mdash;")
s=Replace(s,"¦","&brvbar;")
s=Replace(s,"•","&bull;")
s=Replace(s,"‣","—	")
s=Replace(s,"…","&hellip;")
s=Replace(s,"ˆ","&circ;")
s=Replace(s,"¨","&uml;")
s=Replace(s,"˜","&tilde;")
s=Replace(s,"‹","&lsaquo;")
s=Replace(s,"›","&rsaquo;")
s=Replace(s,"«","&laquo;")
s=Replace(s,"»","&raquo;")
s=Replace(s,"‾","&oline;")
s=Replace(s,"¿","&iquest;")
s=Replace(s,"¡","&iexcl;")
s=Replace(s,"‽","—	")
HTMLEntities_Decode=s
End Function%><script language="javascript" runat="server">
function lambda(f){
if(/^function\s*\([ a-z0-9.$_,]*\)\s*{[\S\s]*}$/gim.test(f)){
eval("f = "+f.replace(/(\r|\n)/g,""));
return f;
} else {
return function(){};
}
}
</script><%Function BinaryToText(ByVal Binary)
BinaryToText=BinaryTostring(Binary)
End Function
Function BinaryToString(ByVal Binary)
If vartype(Binary)=8209 Then
BinaryToString=FastBinaryToString(Binary)
Else
BinaryToString=Binary
End If
End Function
Function StringToBinary(Binary)
StringToBinary=StreamStringToBinary(Binary,"utf-8")
End Function
Function FWXBinaryToString(ByVal vByteArray)
If LenB(vByteArray)>0 Then
Dim oRS
Set oRS=Server.CreateObject("ADODB.Recordset")
oRS.Fields.Append "temp",adLongVarChar,LenB(vByteArray)
oRS.Open
oRS.AddNew
oRS("temp").AppendChunk vByteArray
oRS.Update
FWXBinaryToString=oRS("temp")
oRS.Close
Set oRS=Nothing
Else
FWXBinaryToString=vByteArray
End If
End Function
Function SimpleBinaryToString(Binary)
Dim I,S
For I=1 To LenB(Binary)
S=S&Chr(AscB(MidB(Binary,I,1)))
Next
SimpleBinaryToString=S
End Function
Function FastBinaryToString(Binary)
Dim cl1,cl2,cl3,pl1,pl2,pl3
Dim L
cl1=1
cl2=1
cl3=1
L=LenB(Binary)
Do While cl1<=L
pl3=pl3&Chr(AscB(MidB(Binary,cl1,1)))
cl1=cl1+1
cl3=cl3+1
If cl3>300 Then
pl2=pl2&pl3
pl3=""
cl3=1
cl2=cl2+1
If cl2>200 Then
pl1=pl1&pl2
pl2=""
cl2=1
End If
End If
Loop
FastBinaryToString=pl1&pl2&pl3
End Function
Function RSBinaryToString(xBinary)
Dim Binary
If vartype(xBinary)=8 Then Binary=MultiByteToBinary(xBinary)Else Binary=xBinary
Dim RS,LBinary
Const adLongVarChar=201
Set RS=Server.CreateObject("ADODB.Recordset")
LBinary=LenB(Binary)
If LBinary>0 Then
RS.Fields.Append "mBinary",adLongVarChar,LBinary
RS.Open
RS.AddNew
RS("mBinary").AppendChunk Binary
RS.Update
RSBinaryToString=RS("mBinary")&""
RS.Close
KillRS RS
Set RS=Nothing
Else
RSBinaryToString=""
End If
End Function
Function MultiByteToBinary(MultiByte)
Dim RS,LMultiByte,Binary
Const adLongVarBinary=205
Set RS=Server.CreateObject("ADODB.Recordset")
LMultiByte=LenB(MultiByte)
If LMultiByte>0 Then
RS.Fields.Append "mBinary",adLongVarBinary,LMultiByte
RS.Open
RS.AddNew
RS("mBinary").AppendChunk MultiByte&ChrB(0)
RS.Update
Binary=RS("mBinary").GetChunk(LMultiByte)
End If
MultiByteToBinary=Binary
End Function
Function StreamBinaryToString(xBinary,CharSet)
Const adTypeText=2
Const adTypeBinary=1
Dim Binary
If vartype(xBinary)=8 Then Binary=MultiByteToBinary(xBinary)Else Binary=xBinary
Dim BinaryStream
Set BinaryStream=Server.CreateObject("ADODB.Stream")
BinaryStream.Type=adTypeBinary
BinaryStream.Open
BinaryStream.Write Binary
BinaryStream.Position=0
BinaryStream.Type=adTypeText
If Len(CharSet)>0 Then
BinaryStream.CharSet=CharSet
Else
BinaryStream.CharSet="utf-8"'"us-ascii"
End If
StreamBinaryToString=BinaryStream.ReadText
End Function
Function StreamStringToBinary(Text,CharSet)
Const adTypeText=2
Const adTypeBinary=1
Dim BinaryStream
Set BinaryStream=Server.CreateObject("ADODB.Stream")
BinaryStream.Type=adTypeText
If Len(CharSet)>0 Then
BinaryStream.CharSet=CharSet
Else
BinaryStream.CharSet="utf-8"'us-ascii"
End If
BinaryStream.Open
BinaryStream.WriteText Text
BinaryStream.Position=0
BinaryStream.Type=adTypeBinary
BinaryStream.Position=0
StreamStringToBinary=BinaryStream.Read
End Function
function IsValidUTF8(s)
dim i
dim c
dim n
IsValidUTF8=false
i=1
do while i<=len(s)
c=asc(mid(s,i,1))
if c and &H80 then
n=1
do while i+n<len(s)
if(asc(mid(s,i+n,1))and &HC0)<>&H80 then
exit do
end if
n=n+1
loop
select case n
case 1
exit function
case 2
if(c and &HE0)<>&HC0 then
exit function
end if
case 3
if(c and &HF0)<>&HE0 then
exit function
end if
case 4
if(c and &HF8)<>&HF0 then
exit function
end if
case else
exit function
end select
i=i+n
else
i=i+1
end if
loop
IsValidUTF8=true
end function
Function EncodeUTF8(ByVal s):EncodeUTF8=UTF8Encode(s):End Function
Function UTF8Encode(s)
Dim i,c
i=1
Do While i<=Len(s)
c=asc(mid(s,i,1))
If c>=&H80 Then
s=left(s,i-1)+chr(&HC2+((c And &H40)/ &H40))+chr(c And &HBF)+mid(s,i+1)
i=i+1
End If
i=i+1
Loop
UTF8Encode=s
End Function
Function DecodeUTF8(ByVal s):DecodeUTF8=UTF8Decode(s):End Function
Function UTF8Decode(ByVal s)
s=Replace(s,"+"," ")
If InStr(1,s,"%")>0 Then
Dim re,rm,m,rp,o:o=s
Set re=new RegExp
re.Global=True
re.IgnoreCase=True
re.Pattern="(\%C[23])?\%[0-9A-F]\%?[0-9A-F]"
Set rm=re.Execute(s)
For Each m In rm
rp=m
m=UCase(m)
If Len(m)=3 Or UCase(Left(m,3))="%C2" Then
Else
m=Replace(m,"%C3%8","C")
m=Replace(m,"%C3%9","D")
m=Replace(m,"%C3%A","E")
m=Replace(m,"%C3%B","F")
End If
m=Right(m,2)
s=Replace(s,rp,Chr("&H"&m))
Next
Set re=Nothing
End If
UTF8Decode=s
End Function
Function DoCheck(ByVal b)
If b Then DoCheck=" checked=""true"""
End Function
Function DoSelect(ByVal b)
If b Then DoSelect=" selected=""true"""
End Function
Function IsSelected(ByVal a,ByVal b)
IsSelected=DoSelect((CStr(a)=CStr(b)))
End Function
Function IsChecked(ByVal a,ByVal b)
IsChecked=DoCheck((CStr(a)=CStr(b)))
End Function
Function InList(ByVal a,ByVal b)
InList=IsInList(a,b)
End Function
Function IsInList(ByVal a,ByVal b)
a=","&NList(a)&","'Replace(Replace(","&a&","," ,",","),", ",",")
b=","&NList(b)&","'Replace(Replace(","&b&","," ,",","),", ",",")
IsInList=CBool(InStr(1,a,b)>0)Or CBool(InStr(1,b,a)>0)
End Function
Function IsInStrList(ByVal a,ByVal b)
a=Replace(Replace(","&a&","," ,",","),", ",",")
b=Replace(Replace(","&b&","," ,",","),", ",",")
IsInStrList=CBool(InStr(1,a,b)>0)Or CBool(InStr(1,b,a)>0)
End Function
Function ColItemGet(ByRef c,ByVal n)
On Error Resume Next
ColItemGet=c(n)
End Function
Function ColItemSet(ByRef c,ByVal n,ByVal v)
On Error Resume Next
c(n)=v
End Function
Function ColItem(ByRef c,ByVal n)
On Error Resume Next
ColItem=c(n)
End Function
Function ColItemText(ByRef c,ByVal n)
On Error Resume Next
ColItemText=c(n)&""
End Function
Function ColItemNumber(ByRef c,ByVal n)
On Error Resume Next
ColItemNumber=VOZ(c(n))
End Function
Function TextFromField(ByRef c,ByVal fs):TextFromField=TFF(c,fs):End Function
Function TFF(ByRef c,ByVal fs)
Dim f,o:o=""
For Each f In Split(fs,",")
If o="" Then o=c(f)
Next
TFF=o
End Function
Function NumberFromField(ByRef c,ByVal fs):NumberFromField=VFF(c,fs):End Function
Function NumFromField(ByRef c,ByVal fs):NumFromField=VFF(c,fs):End Function
Function ValFromField(ByRef c,ByVal fs):ValFromField=VFF(c,fs):End Function
Function VFF(ByRef c,ByVal fs)
Dim f,o:o=0
For Each f In Split(fs,",")
If o=0 Then o=VOZ(c(f))
Next
VFF=o
End Function
Function PrintJVars(ByVal v)
PrintJVars=PrintColVars(j,v)
End Function
Function PrintCol(ByRef c):PrintCol=PrintCollection(c):End Function
Function PrintCollection(ByRef c):PrintCollection=PrintVars(c,"*"):End Function
Function PrintVars(ByRef c,ByVal v):PrintVars=PrintCollectionVars(c,v):End Function
Function PrintColVars(ByRef c,ByVal v):PrintColVars=PrintCollectionVars(c,v):End Function
Function PrintColFields(ByRef c,ByVal v):PrintColFields=PrintCollectionVars(c,v):End Function
Function PrintCollectionVars(ByRef c,ByVal f)
Dim i,v
Dim o:Set o=New QuickString
If f<>"" And f<>"*" Then
For Each i In Split(f,",")
If IsObject(c(i))Then
v=TOT(c(i),"[object]")
Else
v=c(i)
End If
o.Add i&": "&v&"<br/>"&vbCRLf
Next
Else
If not IsObject(c)Then
o.Add "Not a collection"
ElseIf IsRecordset(c)Then
o.Add "Collection was a recordset<br/>"&vbCRLf&PrintRS(c)
Else
For Each i In c
If IsObject(c(i))Then
v=TOT(c(i),"[object]")
Else
v=c(i)
End If
o.Add i&": "&v&"<br/>"&vbCRLf
Next
End If
End If
PrintCollectionVars=o
End Function
Function DoTransferCollection(ByRef dSrc,ByRef dDst,ByVal pfx,ByVal sfx)
For Each i In dSrc
If IsObject(dSrc(i))Then
Else
dDst(pfx&i&sfx&"")=dSrc(i)
End If
Next
End Function
Function TransferCollection(ByRef dSrc,ByRef dDst)
TransferCollection=DoTransferCollection(dSrc,dDst,"","")
End Function
Function TransferCollectionWithSufix(ByRef dSrc,ByRef dDst,ByVal sfx)
TransferCollectionWithSufix=DoTransferCollection(dSrc,dDst,"",sfx)
End Function
Function TransferCollectionWithPrefix(ByRef dSrc,ByRef dDst,ByVal pfx)
TransferCollectionWithPrefix=DoTransferCollection(dSrc,dDst,pfx,"")
End Function
Function UpdateCollection(ByRef c,ByRef s)
UpdateCollection=UpdateCollectionWith(c,s,"","")
End Function
Function UpdateCollectionWithSufix(ByRef c,ByRef s,ByVal sfx)
UpdateCollectionWithSufix=UpdateCollectionWith(c,s,"",sfx)
End Function
Function UpdateCollectionWithPrefix(ByRef c,ByRef s,ByVal pfx)
UpdateCollectionWithPrefix=UpdateCollectionWith(c,s,pfx,"")
End Function
Function UpdateCollectionWithObject(ByRef c,ByVal s,ByVal pfx,ByVal sfx)
If Not IsObject(c)Then DBG.FatalError "Call to UpdateCollectionWithObject on invalid variable:<br/>'UpdateCollectionWith(ByRef c'"
Dim iPropCount:iPropCount=0
Dim f,fields,si
If IsObject(s)Then
If IsRecordset(s)Then
Set fields=s.Fields
If Exists(s)Then
For Each f In fields
c(pfx&f.Name&sfx)=f.Value
Next
End If
Else
For Each si In s
If IsObject(s(si))Then
iPropCount=iPropCount+1
Set c(pfx&si&sfx)=s(si)
Else
iPropCount=iPropCount+1
ColItemSet c,pfx&si&sfx,ColItemGet(s,si)
End If
Next
End If
End If
UpdateCollectionWithObject=iPropCount
End Function
Function UpdateCollectionJSON(ByRef c,ByVal s)
UpdateCollectionJSON=UpdateCollectionJSONWithPrefix(c,s,"")
End Function
Function UpdateCollectionJSONWithPrefix(ByRef c,ByVal s,ByVal pfx)
If Not IsObject(c)Then DBG.FatalError "Call to UpdateCollectionJSONWithPrefix on invalid variable:<br/>'UpdateCollectionJSONWithPrefix(ByRef c'"
Dim iPropCount:iPropCount=0
Dim f,fields,si
If IsObject(s)Then
For Each si In s
If IsObject(s(si))Then
iPropCount=iPropCount+1
UpdateCollectionJSONWithPrefix c,s(si),pfx&si&"."
Else
iPropCount=iPropCount+1
ColItemSet c,pfx&si&sfx,ColItemGet(s,si)
End If
Next
End If
UpdateCollectionJSONWithPrefix=iPropCount
End Function
Function UpdateCollectionWithString(ByRef c,ByVal s,ByVal pfx,ByVal sfx)
If Not IsObject(c)Then DBG.FatalError "Call to UpdateCollectionWithString on invalid variable:<br/>'UpdateCollectionWithString(ByRef c'"
Dim iPropCount:iPropCount=0
Dim f,fields,si
Dim iFEquals,sFVar,sFVal
Const PTY_NAME="\b([\w\-\_\.]+)\b"
Const TAG_MARKER="~*~"
Const ESC_MARKER="~@~"
Const PTY_MARKER="~%~"
Const rxESC="\\."
Const rxTAG="<((?!(>)).)+?(>)"
Dim s0,rxColProps,i,aTAG,aESC
rxColProps=PTY_NAME&"("&PTY_MARKER&"((?!"&PTY_MARKER&")[\s\S])*?"&PTY_MARKER&")"
s0=s:original=s
If rxTest(s0,rxESC&"")Then:Set aESC=rxMatches(s0,rxESC&""):s0=rxReplace(s0,rxESC&"",ESC_MARKER&""):End If
If rxTest(s0,rxTAG&"")Then:Set aTAG=rxMatches(s0,rxTAG&""):s0=rxReplace(s0,rxTAG&"",TAG_MARKER&""):End If
s0=rxReplace(s0,PTY_NAME&"\s*=\s*(\""((?!\"")[\s\S])*?\""|\'((?!\')[\s\S])*?\')","$1"&PTY_MARKER&"$2"&PTY_MARKER&"")
s0=rxReplace(s0,"[\""\']?"&PTY_MARKER&"[\""\']?",PTY_MARKER&"")
i=0
Do While InStr(1,s0,TAG_MARKER&"")>0
If i>5 Then
If InStr(1,aTAG(i),TAG_MARKER)Then
DIE "Endless Loop:"&aTAG(i)
End If
End If
s0=Replace(s0,TAG_MARKER&"",aTAG(i),1,1)
i=i+1
Loop
If IsObject(aESC)Then
i=0
Do While InStr(1,s0,ESC_MARKER&"")>0
s0=Replace(s0,ESC_MARKER&"",Right(aESC(i),1),1,1)
i=i+1
Loop
End If
If rxTest(s0,rxColProps)Then
Dim sP,aPs,aParts,mP,sPName
Set aPs=rxMatches(s0,rxColProps)
For Each mP In aPs
sP=TOE(mP.Value)
If InStr(1,sP,PTY_MARKER&"")>0 Then
aParts=Split(sP,PTY_MARKER&"")
If UBound(aParts)>0 Then
sPName=aParts(0)
sFVal=aParts(1)
If sPName<>"" Then
iPropCount=iPropCount+1
c(pfx&sPName&sfx)=FWX_ParsedData(sFVal)
End If
End If
Else
If IsSet(sP)Then
iPropCount=iPropCount+1
c(pfx&sP&sfx)=""
End If
End If
Next
Set aPs=Nothing
ElseIf rxTest(s,"^([^\=\&]+=[^\=\&]*?)(&([^\=\&\n\r]+=[^\=\&\n\r]*?)?)*$")Then
Dim qsPars,qsPair
qsPars=Split(s,"&")
For Each qsPair In qsPars
iFEquals=InStr(1,qsPair,"=")
If GTZ(iFEquals)Then
sFVar=rxReplace(Left(qsPair,iFEquals-1)&"","^[\s\n\r]+|[\s\n\r]+$","")
sFVal=rxReplace(Mid(qsPair,iFEquals+1,Len(qsPair))&"","^[\s\n\r]+|[\s\n\r]+$","")
If sFVar<>"" Then
If InStr(1,sFVar,"%")Or InStr(1,sFVar,"+")Then sFVar=UTF8Decode(sFVar)
If InStr(1,sFVal,"%")Or InStr(1,sFVal,"+")Then sFVal=UTF8Decode(sFVal)
iPropCount=iPropCount+1
c(pfx&sFVar&sfx)=FWX_ParsedData(sFVal)
End If
End If
Next
Else
Dim sF,aFPars,sFLine
s=rxReplace(s,"\s*[\r\n]+\s*",vbCRLf)
aFPars=Split(s,vbCRLf)
For Each sFLine In aFPars
iFEquals=InStr(1,sFLine,"=")
If GTZ(iFEquals)Then
sFVar=TOE(Left(sFLine,iFEquals-1)&"")
sFVal=TOE(Mid(sFLine,iFEquals+1,Len(sFLine))&"")
If sFVar<>"" Then
If InStr(1,sFVar,"%")Or InStr(1,sFVar,"+")Then sFVar=UTF8Decode(sFVar)
If InStr(1,sFVal,"%")Or InStr(1,sFVal,"+")Then sFVal=UTF8Decode(sFVal)
iPropCount=iPropCount+1
c(pfx&sFVar&sfx)=FWX_ParsedData(sFVal)
End If
End If
Next
End If
If Not IsObject(s)Then
s=TOE(s)
If IsNumeric(s)Then s=VOZ(TOE(s))
c(pfx&"args"&sfx)=s
c(pfx&"props"&sfx)=iPropCount
End If
UpdateCollectionWithString=iPropCount
End Function
Function UpdateCollectionWith(ByRef c,ByVal s,ByVal pfx,ByVal sfx)
If Not IsObject(c)Then
DBG.FatalError "Call to UpdateCollectionWith on invalid variable: 'UpdateCollectionWith(ByRef c'"
End If
If IsObject(s)Then
UpdateCollectionWith=UpdateCollectionWithObject(c,s,pfx,sfx)
Else
UpdateCollectionWith=UpdateCollectionWithString(c,s,pfx,sfx)
End If
End Function
Function FilterJ(ByRef s)
Set FilterJ=FilterCol(j,s)
End Function
Function FilterCol(ByVal c,ByVal s)
Set FilterCol=FilterCollection(c,s)
End Function
Function FilterCollection(ByVal c,ByVal s)
Dim o,a,b,f,v,i,j
Set o=Dictionary()
For Each s In Split(s,",")
a="":b="":f="":v=""
If InStr(1,s,"=")Then
a=rxSelect(s,"^[^\=]+")
b=rxSelect(s,"[^\=]+$")
Else
a=s
b=s
End If
If InStr(1,b,":")Then
f=rxSelect(b,"[^\:]+$")
b=rxSelect(b,"^[^\:]+")
End If
If IsObject(c(b))Then
Set o(a)=c(b)
Else
v=c(b)
If v="" Then
If InStr(1,b,",")>0 Then
j=Split(b,",")
For Each i In j
If v="" Then v=c(j(i))
Next
End If
End If
If f<>"" Then
If IsObject(parser)Then
v=parser.Format(v,f)
End If
End If
o(a)=v
End If
Next
Set FilterCollection=o
End Function
Function CopyDic(ByRef c):Set CopyDic=CopyCollection(c):End Function
Function CopyCol(ByRef c):Set CopyCol=CopyCollection(c):End Function
Function CopyDictionary(ByVal c):Set CopyDictionary=CopyCollection(c):End Function
Function CopyCollection(ByRef dSrc)
Dim dDst
Set dDst=Dictionary()
TransferCollection dSrc,dDst
Set CopyCollection=dDst
End Function
Function MakeCol(ByVal s):Set MakeCol=StringToDic(s):End Function
Function MakeDic(ByVal s):Set MakeDic=StringToDic(s):End Function
Function BuildCol(ByVal s):Set BuildCol=StringToDic(s):End Function
Function BuildDic(ByVal s):Set BuildDic=StringToDic(s):End Function
Function ConfigCol(ByVal s):Set ConfigCol=StringToDic(s):End Function
Function ConfigDic(ByVal s):Set ConfigDic=StringToDic(s):End Function
Function Str2Dic(ByVal s):Set Str2Dic=StringToDic(s):End Function
Function StrToDic(ByVal s):Set StrToDic=StringToDic(s):End Function
Function String2Dic(ByVal s):Set String2Dic=StringToDic(s):End Function
Function StringToDic(ByVal s)
Dim d
Set d=Dictionary()
UpdateCollection d,s
Set StringToDic=d
End Function
Function Rs2Qry(ByVal d):Rs2Qry=RSToQuery(d):End Function
Function RsToQRY(ByVal d):RsToQRY=RSToQuery(d):End Function
Function Rs2Query(ByVal d):Rs2Query=RSToQuery(d):End Function
Function RSToQuery(ByVal d)
RSToQuery=Dic2Qry(Rs2Dic(d))
End Function
Function Dic2Qry(ByVal d):Dic2Qry=DicToQuery(d):End Function
Function DicToQry(ByVal d):DicToQry=DicToQuery(d):End Function
Function Dic2Query(ByVal d):Dic2Query=DicToQuery(d):End Function
Function DicToQuery(ByVal d)
Dim s:Set s=New QuickString
Dim f,v
For Each f In d
v=UTF8Encode(d(f))
s.Add f&"="&v
Next
s.Divider="&"
DicToQuery=s&""
End Function
Function Qry2Col(ByVal s):Set Qry2Col=QueryToCol(s):End Function
Function QryToCol(ByVal s):Set QryToCol=QueryToCol(s):End Function
Function Query2Col(ByVal s):Set Query2Col=QueryToCol(s):End Function
Function QueryToCol(ByVal s):Set QueryToCol=QueryToDic(s):End Function
Function Qry2Dic(ByVal s):Set Qry2Dic=QueryToDic(s):End Function
Function QryToDic(ByVal s):Set QryToDic=QueryToDic(s):End Function
Function Query2Dic(ByVal s):Set Query2Dic=QueryToDic(s):End Function
Function QueryToDic(ByVal s)
Dim d
If IsObject(s)Then
Set d=s
Else
Set d=Dictionary()
Dim i,p,v
For Each y In Split(s,"&")
If y<>"" Then
p=y:v="y":i=InStr(1,y,"=")
If i>0 Then
p=Left(y,i-1):v=Mid(y,i+1)
End If
If InStr(1,p,"%")Or InStr(1,p,"+")Then p=UTF8Decode(p)
If InStr(1,v,"%")Or InStr(1,v,"+")Then v=UTF8Decode(v)
If d.Exists(p)Then d(p)=d(p)&", "
d(p)=d(p)&v
End If
Next
End If
Set QueryToDic=d
End Function
Function SortDictionary(objDict)
Dim strDict()
Dim objKey
Dim strKey,strItem
Dim X,Y,Z
Z=objDict.Count
If Z>1 Then
ReDim strDict(Z,2)
X=0
For Each objKey In objDict
strDict(X,1)=CStr(objKey)
strDict(X,2)=CStr(objDict(objKey))
X=X+1
Next
For X=0 To(Z-2)
For Y=X To(Z-1)
If StrComp(strDict(X,1),strDict(Y,1),vbTextCompare)>0 Then
strKey=strDict(X,1)
strItem=strDict(X,2)
strDict(X,1)=strDict(Y,1)
strDict(X,2)=strDict(Y,2)
strDict(Y,1)=strKey
strDict(Y,2)=strItem
End If
Next
Next
objDict.RemoveAll
For X=0 To(Z-1)
objDict.Add strDict(X,1),strDict(X,2)
Next
End If
Set SortDictionary=objDict
End Function
Function DicToJSON(ByRef c):DicToJSON=Col2JS(c):End Function
Function Dic2JSON(ByRef c):Dic2JSON=Col2JS(c):End Function
Function DicToJS(ByRef c):DicToJS=Col2JS(c):End Function
Function Dic2JS(ByRef c):Dic2JS=Col2JS(c):End Function
Function ColToJSON(ByRef c):ColToJSON=Col2JS(c):End Function
Function Col2JSON(ByRef c):Col2JSON=Col2JS(c):End Function
Function ColToJS(ByRef c):ColToJS=Col2JS(c):End Function
Function Col2JS(ByRef c)
Dim o
o=GenerateJson(c)
o=Mid(o,2,Len(o)-2)
Col2JS=o
End Function
Function DIC2XML(ByRef c)
Dim dat
Set dat=New QuickString
dat.Divider=vbCRLf
For Each f In c
If f<>"" Then
Dim v
If IsObject(c(f))Then
v=OTOT(c(f),"[vbscript object]")
Else
v=c(f)&""
End If
If IsNumeric(v)And(Not StartsWith(v,"0"))Then
ElseIf v="True" Or v="False" Then
v=IIf(v,"y","n")
Else
v=Decode(v)
If v<>"" Then
If rxTest(v,"[^\w\s]")Then
v="<![CDATA["&v&"]]>"
End If
End If
End If
dat.Add "<"&f&">"&v&"</"&f&">"
End If
Next
DIC2XML=dat
End Function
Function NumberSplit(ByVal c,ByVal d):NumberSplit=NSplit(c,d):End Function
Function NumbersSplit(ByVal c,ByVal d):NumbersSplit=NSplit(c,d):End Function
Function SplitNumbers(ByVal c,ByVal d):SplitNumbers=NSplit(c,d):End Function
Function NSplit(ByVal c,ByVal d)
Dim x,i:i=0
x=Split(c,d)
Do While i<=UBound(x)
x(i)=VOZ(x(i))
i=i+1
Loop
NSplit=x
End Function
Function NumberList(ByVal s):NumberList=NList(s):End Function
Function NumbersList(ByVal s):NumbersList=NList(s):End Function
Function NList(ByVal s)
s=rxReplace(Trim(Replace(s,","," ")),"\s+",",")
Dim i,d
Set d=Dictionary()
For Each i In NSplit(s,",")
d(i)=i
Next
s=""
For Each i In d
s=s&","&i
Next
s=Mid(s,2)
NList=s
End Function
Function UniqueList(ByVal s)
Dim o,b
Set o=New QuickString
o.Divider=","
For Each b In Split(s,",")
If Not o.Has(b)Then o.Add b
Next
UniqueList=o
End Function
Sub CacheResponseExpires(ByVal intCacheExpiry)
intCacheExpiry=VOZ(intCacheExpiry)
If DBG.Show Then
DBG "CacheResponseExpires("&intCacheExpiry&") ABORTED - can't set headers in debug mode"
Exit Sub
End If
Select Case intCacheExpiry
Case 262080:Response.ExpiresAbsolute=DateAdd("m",6,Now())
Case 525600:Response.ExpiresAbsolute=DateAdd("yyyy",1,Now())
Case Else:Response.ExpiresAbsolute=DateAdd("n",intCacheExpiry,Now())
End Select
End Sub
Sub PublicCacheResponse(ByVal intCacheExpiry)
CacheResponseExpires intCacheExpiry
If DBG.Show Then
DBG "PublicCacheResponse("&intCacheExpiry&") ABORTED - can't set headers in debug mode"
Exit Sub
End If
Response.CacheControl="public"
End Sub
Sub PrivateCacheResponse(ByVal intCacheExpiry)
CacheResponseExpires intCacheExpiry
FWX_AddHeader "Cache-Control","proxy-revalidate"
FWX_AddHeader "Cache-Control","must-revalidate"
FWX_AddHeader "Pragma","no-cache"
If DBG.Show Then
DBG "PrivateCacheResponse("&intCacheExpiry&") ABORTED - can't set headers in debug mode"
Exit Sub
End If
Response.CacheControl="private"
End Sub
Sub CacheResponse(ByVal intCacheExpiry)
PrivateCacheResponse intCacheExpiry
End Sub
Sub CacheControl(ByVal sKEY)
Dim sHI_INM,sHO_ETG
Dim sHI_IMS,sHO_LMD
Dim blnDo304:blnDo304=False
sHI_INM=TOE(Request.ServerVariables("HTTP_IF_NONE_MATCH")&"")
sHI_IMS=TOE(Request.ServerVariables("HTTP_IF_MODIFIED_SINCE")&"")
sHO_LMD=FmtDate(sKEY,fwxDateRFC822)
sHO_ETG="""FWX_"&MD5(sKEY&FWX.Setting("EDITION"))&""""
If sHI_INM<>"" And sHI_IMS<>"" Then
blnDo304=CBool(sHI_INM=sHO_ETG And sHI_IMS=sHO_LMD)
ElseIf sHI_INM<>"" And sHI_IMS="" Then
blnDo304=CBool(sHI_INM=sHO_ETG)
ElseIf sHI_IMS<>"" And sHI_INM="" Then
blnDo304=CBool(sHI_IMS=sHO_LMD)
Else
End If
FWX_AddHeader "ETag",sHO_ETG
If sHO_LMD<>"" Then FWX_AddHeader "Last-Modified",sHO_LMD
If blnDo304 Then
Response.Clear()
FWX_ResponseStatus "304 Not Modified"
Response.Flush()
Response.End()
End If
End Sub
Sub PublicCacheControl(ByVal sKEY,ByVal intCacheExpiry)
PublicCacheResponse intCacheExpiry
CacheControl sKEY
End Sub
Sub PrivateCacheControl(ByVal sKEY,ByVal intCacheExpiry)
PrivateCacheResponse intCacheExpiry
CacheControl sKEY
End Sub%><script language="javascript" runat="server">
function Date_toUtcString(d){
var dDate=new Date(d);
return Math.round(dDate.getTime()/ 1000);
};
</script><%Function UnixDate(ByVal d)
UnixDate=DateToUTCString(d)
End Function
Function UTCStringToDate(ByVal u)
UTCStringToDate=DateAdd("s",u,"01/01/1970 00:00:00")
End Function
Function DateToUTCString(ByVal d)
DateToUTCString=Date_toUtcString(d)
End Function
Function ZuluDate(ByVal d)
d=NewDate(d)
ZuluDate=ConvertDate(Now(),"Zulu")
End Function
Function UTCDate()
UTCDate=Date()
End Function
Function UTCNow()
UTCNow=Now()
End Function
Function ConvertDateFrom(ByVal d,ByVal tzTarget,ByVal tzOrigin)
Dim a,b,k,u,hk,hd,hda,hdb,tzTargetParsed,tzOriginParsed
If d<>"" Then
a=FmtDate(d,"%Y-%M-%D+%H:%N:%S")
If tzTarget="" Then tzTargetParsed="UTC"
If tzTargetParsed="" Then tzTargetParsed=CountryTimezone(tzTarget)
If tzTargetParsed="" Then
DBG.FatalError "Invalid tzTargetParsed = "&tzTarget&"."
End If
If tzOrigin="" Then tzOriginParsed="UTC"
If tzOriginParsed="" Then tzOriginParsed=CountryTimezone(tzOrigin)
If tzOriginParsed="" Then
DBG.FatalError "Invalid tzOriginParsed = "&tzOrigin&"."
End If
k="date="&a&"&origin="&tzOriginParsed&"&target="&tzTargetParsed&"&"
DBG "FWX :: Date :: ConvertDate > k = "&k&""
If tzTargetParsed<>"" And tzTargetParsed<>"-" Then
If LCase(tzTargetParsed)=LCase(tzOriginParsed)Then
b=Replace(a,"+"," ")
Else
b=j("TZ:"&k)
If b="" Then
hk="from:"&tzOriginParsed&" to:"&tzTargetParsed&" on:"&FmtDate(d,"YMD")
hd=j("TZ:"&hk)
If hd="" Then
hd=Application.Contents("TZ:"&hk)
End If
If hd<>"" Then
b=DateAdd("h",hd,Replace(a,"+"," "))
Else
u="/fwx/php/datetime/convert.php?"&k
b=GetResponse(u)
hda=Replace(a,"+"," ")
hdb=Replace(b,"+"," ")
If IsDate(hdb)And IsDate(hda)Then
hd=DateDiff("h",hda,hdb)
j("TZ:"&hk)=hd
If Application.Contents("TZ:"&hk)="" Then
Application.Contents("TZ:"&hk)=hd
End If
End If
End If
j("TZ:"&k)=b
End If
End If
b=FmtDate(b,"%Y-%b-%D %H:%N:%S")
Else
b=a
End If
Else
End If
DBG "FWX :: Date :: ConvertDateFrom ("&d&","&tzTarget&","&tzOrigin&") = b = "&b
ConvertDateFrom=b
End Function
Function ConvertDate(ByVal d,ByVal tzTarget):ConvertDate=UTC2TZ(d,tzTarget):End Function
Function UTCDateInTimezone(ByVal d,ByVal tzTarget):UTCDateInTimezone=UTC2TZ(d,tzTarget):End Function
Function TimezoneDateInUTC(ByVal d,ByVal tzTarget):TimezoneDateInUTC=TZ2UTC(d,tzTarget):End Function
Function UTCtoLocalDate(ByVal d,ByVal tzTarget):UTCtoLocalDate=UTC2TZ(d,tzTarget):End Function
Function LocalDatetoUTC(ByVal d,ByVal tzTarget):LocalDatetoUTC=TZ2UTC(d,tzTarget):End Function
Function UTCtoLocal(ByVal d,ByVal tzTarget):UTCtoLocal=UTC2TZ(d,tzTarget):End Function
Function LocaltoUTC(ByVal d,ByVal tzTarget):LocaltoUTC=TZ2UTC(d,tzTarget):End Function
Function UTCtoTZ(ByVal d,ByVal tzTarget):UTCtoTZ=UTC2TZ(d,tzTarget):End Function
Function TZtoUTC(ByVal d,ByVal tzTarget):TZtoUTC=TZ2UTC(d,tzTarget):End Function
Function DateToTZ(ByVal d,ByVal tzTarget):DateToTZ=UTC2TZ(d,tzTarget):End Function
Function DateToUTC(ByVal d,ByVal tzTarget):DateToUTC=TZ2UTC(d,tzTarget):End Function
Function UTC2TZ(ByVal d,ByVal tzTarget)
UTC2TZ=ConvertDateFrom(d,tzTarget,"UTC")
End Function
Function TZ2UTC(ByVal d,ByVal tzTarget)
TZ2UTC=ConvertDateFrom(d,"UTC",tzTarget)
End Function
Function CountryTimezone(ByVal xx)
Dim o
Select Case UCase(xx)
Case "AD":o="Europe/Andorra"
Case "AE":o="Asia/Dubai"
Case "AF":o="Asia/Kabul"
Case "AG":o="America/Antigua"
Case "AI":o="America/Anguilla"
Case "AL":o="Europe/Tirane"
Case "AM":o="Asia/Yerevan"
Case "AO":o="Africa/Luanda"
Case "AQ":o="Antarctica/McMurdo"
Case "AQ":o="Antarctica/Rothera"
Case "AQ":o="Antarctica/Palmer"
Case "AQ":o="Antarctica/Mawson"
Case "AQ":o="Antarctica/Davis"
Case "AQ":o="Antarctica/Casey"
Case "AQ":o="Antarctica/Vostok"
Case "AQ":o="Antarctica/DumontDUrville"
Case "AQ":o="Antarctica/Syowa"
Case "AQ":o="Antarctica/Troll"
Case "AR":o="America/Argentina/Buenos_Aires"
Case "AR":o="America/Argentina/Cordoba"
Case "AR":o="America/Argentina/Salta"
Case "AR":o="America/Argentina/Jujuy"
Case "AR":o="America/Argentina/Tucuman"
Case "AR":o="America/Argentina/Catamarca"
Case "AR":o="America/Argentina/La_Rioja"
Case "AR":o="America/Argentina/San_Juan"
Case "AR":o="America/Argentina/Mendoza"
Case "AR":o="America/Argentina/San_Luis"
Case "AR":o="America/Argentina/Rio_Gallegos"
Case "AR":o="America/Argentina/Ushuaia"
Case "AS":o="Pacific/Pago_Pago"
Case "AT":o="Europe/Vienna"
Case "AU":o="Australia/Lord_Howe"
Case "AU":o="Antarctica/Macquarie"
Case "AU":o="Australia/Hobart"
Case "AU":o="Australia/Currie"
Case "AU":o="Australia/Melbourne"
Case "AU":o="Australia/Sydney"
Case "AU":o="Australia/Broken_Hill"
Case "AU":o="Australia/Brisbane"
Case "AU":o="Australia/Lindeman"
Case "AU":o="Australia/Adelaide"
Case "AU":o="Australia/Darwin"
Case "AU":o="Australia/Perth"
Case "AU":o="Australia/Eucla"
Case "AW":o="America/Aruba"
Case "AX":o="Europe/Mariehamn"
Case "AZ":o="Asia/Baku"
Case "BA":o="Europe/Sarajevo"
Case "BB":o="America/Barbados"
Case "BD":o="Asia/Dhaka"
Case "BE":o="Europe/Brussels"
Case "BF":o="Africa/Ouagadougou"
Case "BG":o="Europe/Sofia"
Case "BH":o="Asia/Bahrain"
Case "BI":o="Africa/Bujumbura"
Case "BJ":o="Africa/Porto-Novo"
Case "BL":o="America/St_Barthelemy"
Case "BM":o="Atlantic/Bermuda"
Case "BN":o="Asia/Brunei"
Case "BO":o="America/La_Paz"
Case "BQ":o="America/Kralendijk"
Case "BR":o="America/Noronha"
Case "BR":o="America/Belem"
Case "BR":o="America/Fortaleza"
Case "BR":o="America/Recife"
Case "BR":o="America/Araguaina"
Case "BR":o="America/Maceio"
Case "BR":o="America/Bahia"
Case "BR":o="America/Sao_Paulo"
Case "BR":o="America/Campo_Grande"
Case "BR":o="America/Cuiaba"
Case "BR":o="America/Santarem"
Case "BR":o="America/Porto_Velho"
Case "BR":o="America/Boa_Vista"
Case "BR":o="America/Manaus"
Case "BR":o="America/Eirunepe"
Case "BR":o="America/Rio_Branco"
Case "BS":o="America/Nassau"
Case "BT":o="Asia/Thimphu"
Case "BW":o="Africa/Gaborone"
Case "BY":o="Europe/Minsk"
Case "BZ":o="America/Belize"
Case "CA":o="America/St_Johns"
Case "CA":o="America/Halifax"
Case "CA":o="America/Glace_Bay"
Case "CA":o="America/Moncton"
Case "CA":o="America/Goose_Bay"
Case "CA":o="America/Blanc-Sablon"
Case "CA":o="America/Toronto"
Case "CA":o="America/Nipigon"
Case "CA":o="America/Thunder_Bay"
Case "CA":o="America/Iqaluit"
Case "CA":o="America/Pangnirtung"
Case "CA":o="America/Resolute"
Case "CA":o="America/Atikokan"
Case "CA":o="America/Rankin_Inlet"
Case "CA":o="America/Winnipeg"
Case "CA":o="America/Rainy_River"
Case "CA":o="America/Regina"
Case "CA":o="America/Swift_Current"
Case "CA":o="America/Edmonton"
Case "CA":o="America/Cambridge_Bay"
Case "CA":o="America/Yellowknife"
Case "CA":o="America/Inuvik"
Case "CA":o="America/Creston"
Case "CA":o="America/Dawson_Creek"
Case "CA":o="America/Vancouver"
Case "CA":o="America/Whitehorse"
Case "CA":o="America/Dawson"
Case "CC":o="Indian/Cocos"
Case "CD":o="Africa/Kinshasa"
Case "CD":o="Africa/Lubumbashi"
Case "CF":o="Africa/Bangui"
Case "CG":o="Africa/Brazzaville"
Case "CH":o="Europe/Zurich"
Case "CI":o="Africa/Abidjan"
Case "CK":o="Pacific/Rarotonga"
Case "CL":o="America/Santiago"
Case "CL":o="Pacific/Easter"
Case "CM":o="Africa/Douala"
Case "CN":o="Asia/Shanghai"
Case "CN":o="Asia/Urumqi"
Case "CO":o="America/Bogota"
Case "CR":o="America/Costa_Rica"
Case "CU":o="America/Havana"
Case "CV":o="Atlantic/Cape_Verde"
Case "CW":o="America/Curacao"
Case "CX":o="Indian/Christmas"
Case "CY":o="Asia/Nicosia"
Case "CZ":o="Europe/Prague"
Case "DE":o="Europe/Berlin"
Case "DE":o="Europe/Busingen"
Case "DJ":o="Africa/Djibouti"
Case "DK":o="Europe/Copenhagen"
Case "DM":o="America/Dominica"
Case "DO":o="America/Santo_Domingo"
Case "DZ":o="Africa/Algiers"
Case "EC":o="America/Guayaquil"
Case "EC":o="Pacific/Galapagos"
Case "EE":o="Europe/Tallinn"
Case "EG":o="Africa/Cairo"
Case "EH":o="Africa/El_Aaiun"
Case "ER":o="Africa/Asmara"
Case "ES":o="Europe/Madrid"
Case "ES":o="Africa/Ceuta"
Case "ES":o="Atlantic/Canary"
Case "ET":o="Africa/Addis_Ababa"
Case "FI":o="Europe/Helsinki"
Case "FJ":o="Pacific/Fiji"
Case "FK":o="Atlantic/Stanley"
Case "FM":o="Pacific/Chuuk"
Case "FM":o="Pacific/Pohnpei"
Case "FM":o="Pacific/Kosrae"
Case "FO":o="Atlantic/Faroe"
Case "FR":o="Europe/Paris"
Case "GA":o="Africa/Libreville"
Case "GB":o="Europe/London"
Case "UK":o="Europe/London"
Case "GD":o="America/Grenada"
Case "GE":o="Asia/Tbilisi"
Case "GF":o="America/Cayenne"
Case "GG":o="Europe/Guernsey"
Case "GH":o="Africa/Accra"
Case "GI":o="Europe/Gibraltar"
Case "GL":o="America/Godthab"
Case "GL":o="America/Danmarkshavn"
Case "GL":o="America/Scoresbysund"
Case "GL":o="America/Thule"
Case "GM":o="Africa/Banjul"
Case "GN":o="Africa/Conakry"
Case "GP":o="America/Guadeloupe"
Case "GQ":o="Africa/Malabo"
Case "GR":o="Europe/Athens"
Case "GS":o="Atlantic/South_Georgia"
Case "GT":o="America/Guatemala"
Case "GU":o="Pacific/Guam"
Case "GW":o="Africa/Bissau"
Case "GY":o="America/Guyana"
Case "HK":o="Asia/Hong_Kong"
Case "HN":o="America/Tegucigalpa"
Case "HR":o="Europe/Zagreb"
Case "HT":o="America/Port-au-Prince"
Case "HU":o="Europe/Budapest"
Case "ID":o="Asia/Jakarta"
Case "ID":o="Asia/Pontianak"
Case "ID":o="Asia/Makassar"
Case "ID":o="Asia/Jayapura"
Case "IE":o="Europe/Dublin"
Case "IL":o="Asia/Jerusalem"
Case "IM":o="Europe/Isle_of_Man"
Case "IN":o="Asia/Kolkata"
Case "IO":o="Indian/Chagos"
Case "IQ":o="Asia/Baghdad"
Case "IR":o="Asia/Tehran"
Case "IS":o="Atlantic/Reykjavik"
Case "IT":o="Europe/Rome"
Case "JE":o="Europe/Jersey"
Case "JM":o="America/Jamaica"
Case "JO":o="Asia/Amman"
Case "JP":o="Asia/Tokyo"
Case "KE":o="Africa/Nairobi"
Case "KG":o="Asia/Bishkek"
Case "KH":o="Asia/Phnom_Penh"
Case "KI":o="Pacific/Tarawa"
Case "KI":o="Pacific/Enderbury"
Case "KI":o="Pacific/Kiritimati"
Case "KM":o="Indian/Comoro"
Case "KN":o="America/St_Kitts"
Case "KP":o="Asia/Pyongyang"
Case "KR":o="Asia/Seoul"
Case "KW":o="Asia/Kuwait"
Case "KY":o="America/Cayman"
Case "KZ":o="Asia/Almaty"
Case "KZ":o="Asia/Qyzylorda"
Case "KZ":o="Asia/Aqtobe"
Case "KZ":o="Asia/Aqtau"
Case "KZ":o="Asia/Oral"
Case "LA":o="Asia/Vientiane"
Case "LB":o="Asia/Beirut"
Case "LC":o="America/St_Lucia"
Case "LI":o="Europe/Vaduz"
Case "LK":o="Asia/Colombo"
Case "LR":o="Africa/Monrovia"
Case "LS":o="Africa/Maseru"
Case "LT":o="Europe/Vilnius"
Case "LU":o="Europe/Luxembourg"
Case "LV":o="Europe/Riga"
Case "LY":o="Africa/Tripoli"
Case "MA":o="Africa/Casablanca"
Case "MC":o="Europe/Monaco"
Case "MD":o="Europe/Chisinau"
Case "ME":o="Europe/Podgorica"
Case "MF":o="America/Marigot"
Case "MG":o="Indian/Antananarivo"
Case "MH":o="Pacific/Majuro"
Case "MH":o="Pacific/Kwajalein"
Case "MK":o="Europe/Skopje"
Case "ML":o="Africa/Bamako"
Case "MM":o="Asia/Rangoon"
Case "MN":o="Asia/Ulaanbaatar"
Case "MN":o="Asia/Hovd"
Case "MN":o="Asia/Choibalsan"
Case "MO":o="Asia/Macau"
Case "MP":o="Pacific/Saipan"
Case "MQ":o="America/Martinique"
Case "MR":o="Africa/Nouakchott"
Case "MS":o="America/Montserrat"
Case "MT":o="Europe/Malta"
Case "MU":o="Indian/Mauritius"
Case "MV":o="Indian/Maldives"
Case "MW":o="Africa/Blantyre"
Case "MX":o="America/Mexico_City"
Case "MX":o="America/Cancun"
Case "MX":o="America/Merida"
Case "MX":o="America/Monterrey"
Case "MX":o="America/Matamoros"
Case "MX":o="America/Mazatlan"
Case "MX":o="America/Chihuahua"
Case "MX":o="America/Ojinaga"
Case "MX":o="America/Hermosillo"
Case "MX":o="America/Tijuana"
Case "MX":o="America/Santa_Isabel"
Case "MX":o="America/Bahia_Banderas"
Case "MY":o="Asia/Kuala_Lumpur"
Case "MY":o="Asia/Kuching"
Case "MZ":o="Africa/Maputo"
Case "NA":o="Africa/Windhoek"
Case "NC":o="Pacific/Noumea"
Case "NE":o="Africa/Niamey"
Case "NF":o="Pacific/Norfolk"
Case "NG":o="Africa/Lagos"
Case "NI":o="America/Managua"
Case "NL":o="Europe/Amsterdam"
Case "NO":o="Europe/Oslo"
Case "NP":o="Asia/Kathmandu"
Case "NR":o="Pacific/Nauru"
Case "NU":o="Pacific/Niue"
Case "NZ":o="Pacific/Auckland"
Case "NZ":o="Pacific/Chatham"
Case "OM":o="Asia/Muscat"
Case "PA":o="America/Panama"
Case "PE":o="America/Lima"
Case "PF":o="Pacific/Tahiti"
Case "PF":o="Pacific/Marquesas"
Case "PF":o="Pacific/Gambier"
Case "PG":o="Pacific/Port_Moresby"
Case "PG":o="Pacific/Bougainville"
Case "PH":o="Asia/Manila"
Case "PK":o="Asia/Karachi"
Case "PL":o="Europe/Warsaw"
Case "PM":o="America/Miquelon"
Case "PN":o="Pacific/Pitcairn"
Case "PR":o="America/Puerto_Rico"
Case "PS":o="Asia/Gaza"
Case "PS":o="Asia/Hebron"
Case "PT":o="Europe/Lisbon"
Case "PT":o="Atlantic/Madeira"
Case "PT":o="Atlantic/Azores"
Case "PW":o="Pacific/Palau"
Case "PY":o="America/Asuncion"
Case "QA":o="Asia/Qatar"
Case "RE":o="Indian/Reunion"
Case "RO":o="Europe/Bucharest"
Case "RS":o="Europe/Belgrade"
Case "RU":o="Europe/Kaliningrad"
Case "RU":o="Europe/Moscow"
Case "RU":o="Europe/Simferopol"
Case "RU":o="Europe/Volgograd"
Case "RU":o="Europe/Samara"
Case "RU":o="Asia/Yekaterinburg"
Case "RU":o="Asia/Omsk"
Case "RU":o="Asia/Novosibirsk"
Case "RU":o="Asia/Novokuznetsk"
Case "RU":o="Asia/Krasnoyarsk"
Case "RU":o="Asia/Irkutsk"
Case "RU":o="Asia/Chita"
Case "RU":o="Asia/Yakutsk"
Case "RU":o="Asia/Khandyga"
Case "RU":o="Asia/Vladivostok"
Case "RU":o="Asia/Sakhalin"
Case "RU":o="Asia/Ust-Nera"
Case "RU":o="Asia/Magadan"
Case "RU":o="Asia/Srednekolymsk"
Case "RU":o="Asia/Kamchatka"
Case "RU":o="Asia/Anadyr"
Case "RW":o="Africa/Kigali"
Case "SA":o="Asia/Riyadh"
Case "SB":o="Pacific/Guadalcanal"
Case "SC":o="Indian/Mahe"
Case "SD":o="Africa/Khartoum"
Case "SE":o="Europe/Stockholm"
Case "SG":o="Asia/Singapore"
Case "SH":o="Atlantic/St_Helena"
Case "SI":o="Europe/Ljubljana"
Case "SJ":o="Arctic/Longyearbyen"
Case "SK":o="Europe/Bratislava"
Case "SL":o="Africa/Freetown"
Case "SM":o="Europe/San_Marino"
Case "SN":o="Africa/Dakar"
Case "SO":o="Africa/Mogadishu"
Case "SR":o="America/Paramaribo"
Case "SS":o="Africa/Juba"
Case "ST":o="Africa/Sao_Tome"
Case "SV":o="America/El_Salvador"
Case "SX":o="America/Lower_Princes"
Case "SY":o="Asia/Damascus"
Case "SZ":o="Africa/Mbabane"
Case "TC":o="America/Grand_Turk"
Case "TD":o="Africa/Ndjamena"
Case "TF":o="Indian/Kerguelen"
Case "TG":o="Africa/Lome"
Case "TH":o="Asia/Bangkok"
Case "TJ":o="Asia/Dushanbe"
Case "TK":o="Pacific/Fakaofo"
Case "TL":o="Asia/Dili"
Case "TM":o="Asia/Ashgabat"
Case "TN":o="Africa/Tunis"
Case "TO":o="Pacific/Tongatapu"
Case "TR":o="Europe/Istanbul"
Case "TT":o="America/Port_of_Spain"
Case "TV":o="Pacific/Funafuti"
Case "TW":o="Asia/Taipei"
Case "TZ":o="Africa/Dar_es_Salaam"
Case "UA":o="Europe/Kiev"
Case "UA":o="Europe/Uzhgorod"
Case "UA":o="Europe/Zaporozhye"
Case "UG":o="Africa/Kampala"
Case "UM":o="Pacific/Johnston"
Case "UM":o="Pacific/Midway"
Case "UM":o="Pacific/Wake"
Case "US":o="America/New_York"
Case "US":o="America/Detroit"
Case "US":o="America/Kentucky/Louisville"
Case "US":o="America/Kentucky/Monticello"
Case "US":o="America/Indiana/Indianapolis"
Case "US":o="America/Indiana/Vincennes"
Case "US":o="America/Indiana/Winamac"
Case "US":o="America/Indiana/Marengo"
Case "US":o="America/Indiana/Petersburg"
Case "US":o="America/Indiana/Vevay"
Case "US":o="America/Chicago"
Case "US":o="America/Indiana/Tell_City"
Case "US":o="America/Indiana/Knox"
Case "US":o="America/Menominee"
Case "US":o="America/North_Dakota/Center"
Case "US":o="America/North_Dakota/New_Salem"
Case "US":o="America/North_Dakota/Beulah"
Case "US":o="America/Denver"
Case "US":o="America/Boise"
Case "US":o="America/Phoenix"
Case "US":o="America/Los_Angeles"
Case "US":o="America/Metlakatla"
Case "US":o="America/Anchorage"
Case "US":o="America/Juneau"
Case "US":o="America/Sitka"
Case "US":o="America/Yakutat"
Case "US":o="America/Nome"
Case "US":o="America/Adak"
Case "US":o="Pacific/Honolulu"
Case "UY":o="America/Montevideo"
Case "UZ":o="Asia/Samarkand"
Case "UZ":o="Asia/Tashkent"
Case "VA":o="Europe/Vatican"
Case "VC":o="America/St_Vincent"
Case "VE":o="America/Caracas"
Case "VG":o="America/Tortola"
Case "VI":o="America/St_Thomas"
Case "VN":o="Asia/Ho_Chi_Minh"
Case "VU":o="Pacific/Efate"
Case "WF":o="Pacific/Wallis"
Case "WS":o="Pacific/Apia"
Case "YE":o="Asia/Aden"
Case "YT":o="Indian/Mayotte"
Case "ZA":o="Africa/Johannesburg"
Case "ZM":o="Africa/Lusaka"
Case "ZW":o="Africa/Harare"
Case Else:o=xx
End Select
CountryTimezone=o
End Function
Function NewDateReplaceAt(ByVal s)
On Error Resume Next
NewDateReplaceAt=Replace(s,"@"," ")
End Function
Dim NewDateTimezone
Dim ZULU_NewDateToday,ZULU_NewDateNow
Function NewDate(ByVal s)
Dim o
Dim NewDateTimezone
Dim NewDateNow,NewDateToday
s=NewDateReplaceAt(s)
If IsDate(s)Then
o=CDate(s)
Else
If ZULU_NewDateNow="" Then
ZULU_NewDateNow=TOT(UTCNow(),Now())
ZULU_NewDateToday=CDate(DoDate(ZULU_NewDateNow))
End If
NewDateNow=ZULU_NewDateNow
NewDateTimezone=MyTimezone
NewDateToday=CDate(DoDate(NewDateNow))
DBG "FWX.NewDate > ZULU_NewDateNow="&ZULU_NewDateNow
DBG "FWX.NewDate > NewDateTimezone="&NewDateTimezone
DBG "FWX.NewDate > NewDateToday="&NewDateToday
DBG "FWX.NewDate > NewDateNow="&NewDateNow
DBG "FWX.NewDate > s="&s
o=ParseDateString(s,NewDateNow)
DBG "FWX.NewDate > s("&s&") = o ="&o
End If
NewDate=o
End Function
Function NewLocalDate(ByVal s)
Dim o
Dim NewDateTimezone
Dim NewDateNow,NewDateToday
s=NewDateReplaceAt(s)
If IsDate(s)Then
o=CDate(s)
Else
If ZULU_NewDateNow="" Then
ZULU_NewDateNow=TOT(UTCNow(),Now())
ZULU_NewDateToday=CDate(DoDate(ZULU_NewDateNow))
End If
NewDateNow=ZULU_NewDateNow
NewDateTimezone=MyTimezone
If s<>"zulu" And NewDateTimezone<>"" And NewDateTimezone<>"-" Then
If j(NewDateTimezone&"_Now")="" Then
j(NewDateTimezone&"_Now")=ConvertDate(ZULU_NewDateNow,NewDateTimezone)
End If
NewDateNow=j(NewDateTimezone&"_Now")
End If
NewDateToday=CDate(DoDate(NewDateNow))
DBG "FWX.NewLocalDate > ZULU_NewDateNow="&ZULU_NewDateNow
DBG "FWX.NewLocalDate > NewDateTimezone="&NewDateTimezone
DBG "FWX.NewLocalDate > NewDateToday="&NewDateToday
DBG "FWX.NewLocalDate > NewDateNow="&NewDateNow
DBG "FWX.NewLocalDate > s="&s
o=ParseDateString(s,NewDateNow)
DBG "FWX.NewLocalDate > s("&s&") = o ="&o
End If
NewLocalDate=o
End Function
Function GetDate(ByVal s)
GetDate=NewDate(TOT(s,Date()))
End Function
Function DateFromString(ByVal s):DateFromString=NewDate(s):End Function
Function MakeDateFromString(ByVal s):MakeDateFromString=NewDate(s):End Function
Function MakeDateString(ByVal s):MakeDateString=NewDate(s):End Function
Function YYYY_MM_DD(ByVal s):YYYY_MM_DD=FmtDate(s,"%Y-%M-%D"):End Function
Function YYYYMMDD(ByVal s):YYYYMMDD=FmtDate(s,"%Y%M%D"):End Function
Function YYYYMMDDHHNN(ByVal s):YYYYMMDDHHNN=FmtDate(s,"%Y%M%D%H%N"):End Function
Function ParseDateString(ByVal s,ByVal n)
Dim o,d
d=CDate(DoDate(n))
If IsDate(s)Then
o=s
Else
If rxTest(s,"^[\-0-9]+$")Then
If Len(s)>10 Then
o=DateAdd("s",VOZ(s),"1 Jan 1970 00:00")
ElseIf Abs(VOZ(o))<=365 Then
o=DateAdd("d",VOZ(s),d)
Else
DBG.Warning "Invalid Date Offset"
End If
Else
Select Case LCase(s)
Case "now":o=n
Case "utc","zulu":o=n
Case "today","2day":o=CDate(DateAdd("d",0-0,d))
Case "yesterday","yday":o=CDate(DateAdd("d",0-1,d))
Case "tomorrow","2moro":o=CDate(DateAdd("d",0+1,d))
Case "lws","last-week-start":o=weekstart(DateAdd("ww",0-1,d))
Case "lwe","last-week-end":o=weekend(DateAdd("ww",0-1,d))
Case "lms","last-month-start":o=monthstart(DateAdd("m",0-1,d))
Case "lme","last-month-end":o=monthend(DateAdd("m",0-1,d))
Case "lys","last-year-start":o=yearstart(DateAdd("yyyy",0-1,d))
Case "lye","last-year-end":o=yearend(DateAdd("yyyy",0-1,d))
Case "tws","this-week-start":o=weekstart(DateAdd("ww",0-0,d))
Case "twe","this-week-end":o=weekend(DateAdd("ww",0-0,d))
Case "tms","this-month-start":o=monthstart(DateAdd("m",0-0,d))
Case "tme","this-month-end":o=monthend(DateAdd("m",0-0,d))
Case "tys","this-year-start":o=yearstart(DateAdd("yyyy",0-0,d))
Case "tye","this-year-end":o=yearend(DateAdd("yyyy",0-0,d))
Case "nws","next-week-start":o=weekstart(DateAdd("ww",0+1,d))
Case "nwe","next-week-end":o=weekend(DateAdd("ww",0+1,d))
Case "nms","next-month-start":o=monthstart(DateAdd("m",0+1,d))
Case "nme","next-month-end":o=monthend(DateAdd("m",0+1,d))
Case "nys","next-year-start":o=yearstart(DateAdd("yyyy",0+1,d))
Case "nye","next-year-end":o=yearend(DateAdd("yyyy",0+1,d))
Case "ws","week-start":o=weekstart(DateAdd("ww",0-0,d))
Case "we","week-end":o=weekend(DateAdd("ww",0-0,d))
Case "ms","month-start":o=monthstart(DateAdd("m",0-0,d))
Case "me","month-end":o=monthend(DateAdd("m",0-0,d))
Case "ys","year-start":o=yearstart(DateAdd("yyyy",0-0,d))
Case "ye","year-end":o=yearend(DateAdd("yyyy",0-0,d))
Case Else
o=DoDate(s)
End Select
End If
End If
ParseDateString=o
End Function
Function DOD(ByVal a,ByVal b)
Err.Clear():On Error Resume Next
Dim s:s=""
If Not s<>"" Then s=NewDate(a&"")
If Not s<>"" Then s=NewDate(b&"")
DOD=s:Err.Clear()
End Function
Function DON(ByVal s):DON=DOD(s,Now()):End Function
Function MakeDate(ByVal iDay,ByVal iMonth,ByVal iYear)
If Not iDay>0 Then iDay=Day(Date())
If Not iMonth>0 Then iMonth=Month(Date())
If Not iYear>0 Then iYear=Year(Date())
MakeDate=CDate(Digits(iDay,2)&"-"&MonthName(iMonth)&"-"&iYear)
End Function
Function WeekStart(ByVal dDate)
iWeekday=Weekday(dDate)
dtOutput=DateAdd("d",1-iWeekday,dDate)
WeekStart=dtOutput
End Function
Function WeekEnd(ByVal dDate)
iWeekday=Weekday(dDate)
dtOutput=DateAdd("d",7-iWeekday,dDate)
WeekEnd=dtOutput
End Function
Function MonthStart(ByVal dDate)
dtOutput=MakeDate(1,Month(dDate),Year(dDate))
MonthStart=dtOutput
End Function
Function MonthEnd(ByVal dDate)
dtOutput=dDate
dtOutput=DateAdd("m",1,dtOutput)
dtOutput=MonthStart(dtOutput)
dtOutput=DateAdd("d",-1,dtOutput)
MonthEnd=dtOutput
End Function
Function YearStart(ByVal dDate)
YearStart=MakeDate(1,1,Year(dDate))
End Function
Function YearEnd(ByVal dDate)
YearEnd=MakeDate(31,12,Year(dDate))
End Function
Function Week(ByVal dDate)
Week=DateDiff("ww",YearStart(Now()),dDate,vbSunday,vbFirstFourDays)
End Function
Function AddWeek(ByVal iChange)
AddWeek=DateAdd("ww",iChange,Date())
End Function
Function NextWeek()
NextWeek=AddWeek(1)
End Function
Function LastWeek()
LastWeek=AddWeek(-1)
End Function
Function Yesterday()
Yesterday=AddDay(-1)
End Function
Function Today()
Today=AddDay(0)
End Function
Function Tomorrow()
Tomorrow=AddDay(1)
End Function
Function AddDay(ByVal iChange)
AddDay=DateAdd("d",iChange,Date())
End Function
Function AddMonth(ByVal iChange)
AddMonth=DateAdd("m",iChange,Date())
End Function
Function NextMonth()
NextMonth=AddMonth(1)
End Function
Function LastMonth()
LastMonth=AddMonth(-1)
End Function
Function AddYear(ByVal iChange)
AddYear=DateAdd("yyyy",iChange,Date())
End Function
Function NextYear()
NextYear=AddYear(1)
End Function
Function LastYear()
LastYear=AddYear(-1)
End Function
Dim G_sDateRange_Options
Function DateRange_Options_For(ByVal sPrefix)
sPrefix=sPrefix
If j("_DateRange_Options"&sPrefix&"")<>"" Then
Else
If Not j("_"&sPrefix&"dates")<>"" Then
j("_"&sPrefix&"dates")=ReloadVars(sPrefix&"start&"&sPrefix&"end&"&sPrefix&"date&")
j("_"&sPrefix&"dates")=Mid(j("_"&sPrefix&"dates"),11)
j("_"&sPrefix&"dates")=Replace(j("_"&sPrefix&"dates"),"%2D","-")
If Not InStr(1,j("_"&sPrefix&"dates"),"-")>0 Then
j("_"&sPrefix&"dates")="&nodates&"
End If
End If
If G_sDateRange_Options<>"" Then
Else
Dim lclString:Set lclString=New QuickString
With lclString
.Divider=vbCRLf
.Add "<option value=''></option>"
.Add "<option value='|prefix|start=&|prefix|end=&|prefix|date=&' style='color:#f00;text-decoration:underline;' title='ever'>all time</option>"
.Add "<option value=''></option>"
.Add "<optgroup label='Daily...'>"
.Add "<option value='|prefix|start=&|prefix|end=&|prefix|date="&DoDate(Date()-1)&"&' style='color:#666;' title='yesterday'>yesterday</option>"
.Add "<option value='|prefix|start=&|prefix|end=&|prefix|date="&DoDate(Date())&"&' style='color:#090;' title='today'>today</option>"
.Add "<option value='|prefix|start=&|prefix|end=&|prefix|date="&DoDate(Date()+1)&"&' style='color:#900;' title='tomorrow'>tomorrow</option>"
.Add "<option value='|prefix|start=&|prefix|end=&|prefix|date="&DoDate(Date()+2)&"&'>"&LCase((WeekdayName(Weekday(Date()+2))))&"</option>"
.Add "<option value='|prefix|start=&|prefix|end=&|prefix|date="&DoDate(Date()+3)&"&'>"&LCase((WeekdayName(Weekday(Date()+3))))&"</option>"
.Add "<option value='|prefix|start=&|prefix|end=&|prefix|date="&DoDate(Date()+4)&"&'>"&LCase((WeekdayName(Weekday(Date()+4))))&"</option>"
.Add "</optgroup>"
.Add "<optgroup label='Weekly...'>"
.Add "<option value='|prefix|start="&DoDate(Date()-7)&"&|prefix|end="&DoDate(Date())&"&' style='color:#069;' title='last7days'>last 7 days</option>"
.Add "<option value='|prefix|start="&DoDate(Date())&"&|prefix|end="&DoDate(Date()+7)&"&|prefix|date=&' style='color:#069;' title='next7days'>next 7 days</option>"
.Add "</optgroup>"
.Add "<optgroup label='Other...'>"
.Add "<option value='|prefix|start="&DoDate(Date())&"&|prefix|end=&|prefix|date=&' style='color:#c33;' title='fromtoday'>from today</option>"
.Add "<option value='|prefix|start=&|prefix|end="&DoDate(Date())&"&|prefix|date=&' style='color:#c33;' title='untiltoday'>until today</option>"
.Add "</optgroup>"
End With
G_sDateRange_Options=lclString
End If
lclString=G_sDateRange_Options
lclString=Replace(lclString,"|prefix|",sPrefix&"")
lclString=Replace(lclString,"value='"&j("_"&sPrefix&"dates")&"'","value='"&j("_"&sPrefix&"dates")&"' SELECTED")
j("_DateRange_Options"&sPrefix&"")=lclString
End If
DateRange_Options_For=j("_DateRange_Options"&sPrefix&"")
End Function
Function DateRange_Options()
DateRange_Options=DateRange_Options_For("")
End Function
Const RedirectsDisabled=False
Function DoRedirect(ByVal s,ByVal sTarget)
If s=Query.FULLURL Then
DBG.FatalError "Redirect loop to: "&s
Exit Function
End If
If s="-" Then s=""
If s="" Then Exit Function
s=s&IIf(rxTest(s,"\?$")Or rxTest(s,"\&$"),"",IIf(rxTest(s,"\?"),"&","?"))
s=rxReplace(s,"\?$","")
If Not IsPost Then
If IsAjax Then s=s&IIf(rxTest(s,"\?"),IIf(rxTest(s,"&(amp;)?$"),"","&"),"?")&"-ajax=y&passed-on"
DBG "Redirect requested to: "&s
DBG ""
If Not DBG.Show Then
With Response
.Clear()
.Status=302
.Expires=0
.ExpiresAbsolute=Date()-1
End With
Response.Redirect s
End If
If Err.Description<>"" Or Err.Number<>0 Then
Err.Clear()
End If
RE
End If
If IsAjax Then
j("redirect")=s
DIEvars "redirect"
End If
RW "<form action='"&s&"' name='RedirectForm' id='RedirectForm' target='"&sTarget&"' method='post' style='"
RW "font-family:arial; width:320px; margin:50px auto; text-align:center;"
RW "'>"
If j("shed-post-data")="" Then RW " "&ReloadFields(ReAssignVars(FormVars(),"~_&"))
RW " <span style='color:#999;font-weight:normal;font-size:24px;'>Loading...</span>"
RW " <br/>"
RW " <input id='RedirecButton' style='"
RW "color:#777;font-size:10px;background:none;border:none;padding:0;margin:0;text-indent:0;text-align:left;"
RW "' type='submit' value='Click here if nothing happens after a few seconds'/>"
RW "</form>"
RW "<scr"&"ipt type='text/javascript'>"&vbCRLf
RW "window.onload=function(){"&vbCRLf
RW " try{"&vbCRLf
RW "  document.getElementById('RedirectForm').submit();"&vbCRLf
RW " }"&vbCRLf
RW " catch(e){"&vbCRLf
RW "  document.getElementById('RedirecButton').click();"&vbCRLf
RW " }"&vbCRLf
RW "};"&vbCRLf
RW "</scr"&"ipt>"
RE
End Function
Function FreshRedirect(ByVal s)
j("shed-post-data")="y"
DoRedirect s&"",""
End Function
Function Redirect(ByVal s)
DoRedirect s&"",""
End Function
Function TopRedirect(ByVal s)
DoRedirect s&"","_top"
End Function
Function PermanentRedirectCache(ByVal f,ByVal t)
If f<>t Then Application.Contents("R:"&f)=t
End Function
Function PermanentRedirectAnticipate(ByVal s)
Dim f,t,x
For Each f In Application.Contents
If StartsWith(f,"R:")Then
t=Application.Contents(f)
f=Mid(f,3)
If InStr(1,s,f)>0 And f<>t Then
x=s
s=rxReplace(s,_
"(['""'])(https?:)?(//("&rxEscape(SERVER_NAME)&"|"&rxEscape(MyWWWDomain)&"))?"&rxEscape(rxReplace(f,"/+$",""))&"/?(['""'])","$1$2$3"&t&"$5")
If s<>x Then
DBG "### PermanentRedirectCache >>>>> ANTICIPATED REDIRECT!!! "&f&" with "&f '&"==="&Encode(s)
End If
End If
End If
Next
PermanentRedirectAnticipate=s
End Function
Function PermanentRedirect(ByVal s)
DBG "REDIRECT REQUESTED to '"&s&"'"
DBG ""
If s="-" Then s=""
If s="" Then Exit Function
s=rxReplace(s,"(index|default)\.(aspx?|html?)$","")
If s="" Then Exit Function
If IsObject(Query)Then
If s=Query.URI Then DBG.FatalError "Infinite Redirect Loop: "&s
Else
If s=QUERY_URI Then DBG.FatalError "Infinite Redirect Loop: "&s
End If
If StartsWith(s,"/")Then
s=rxReplace(s,"^([^\?]+)/$","$1")
s=rxReplace(s,"^([^\?]+)/\?","$1?")
End If
If IsObject(FWX)Then WantsHTTPS=IsTrue(FWX.Setting("https"))
If IsHTTPS Or WantsHTTPS Then
If rxTest(s,"^\/[^\/]")Then
s="https://"&SERVER_NAME&s
End If
End If
If IsAjax Or IsPost Or IsSOAP Then
ElseIf(IsLocal Or RedirectsDisabled Or DBG.Show)And(EXT="" Or EXT="htm" Or EXT="html")Then
RHTM
If IsLocal Then
RW "<p style='font-family:courier new'>"
If IsObject(Query)Then
RW "Leaving<br/><a href='"&Query.RAWURL&"'>"&Query.RAWURL&"</a><br/>"
ElseIf RAWURL<>"" Then
RW "Leaving<br/><a href='"&RAWURL&"'>"&RAWURL&"</a><br/>"
End If
RW "</p>"
End If
RW "<p style='font-family:courier new'>"
RW "<strong>Permanently</strong> Redirecting to<br/><a href='"&s&"'>"&s&"</a><br/>"
RW "</p>"
If RedirectsDisabled Or DBG.Show Then
RW "<p style='color:red;'>ps.: automatic redirects are disabled</p>"
Else
RW "<scr"&"ipt>top.location='"&s&"'</scr"&"ipt>"
End If
RE
End If
FWX_ResponseStatus "301 Moved Permanently"
FWX_AddHeader "Location",s
RE
End Function
Function RedirectForGood(ByVal s)
Call PermanentRedirect(s)
End Function
Function PRedirect(ByVal s)
Call PermanentRedirect(s)
End Function
Function GrossVAT()
GrossVAT=((TAX_RATE/100)/(1+(TAX_RATE/100)))
End Function
Function Gross2VAT():Gross2VAT=GrossVAT:End Function
Function GrossToVAT():GrossToVAT=GrossVAT:End Function
Function GrossNET()
GrossNET=(1/(1+(TAX_RATE/100)))
End Function
Function Gross2NET():Gross2NET=GrossNET:End Function
Function GrossToNET():GrossToNET=GrossNET:End Function
Function NetVAT()
NetVAT=(TAX_RATE/100)
End Function
Function Net2VAT():Net2VAT=NetVAT:End Function
Function NetToVAT():NetToVAT=NetVAT:End Function
Function NetGROSS()
NetGROSS=(1+(TAX_RATE/100))
End Function
Function Net2GROSS():Net2GROSS=NetGROSS:End Function
Function NetToGROSS():NetToGROSS=NetGROSS:End Function
Function FileIcon(ByVal s)
FileIcon=App.Image("files/"&TOT(Extension(s),"file")&"")
End Function
Function Extension(ByVal s)
s=LCase(s&"")
Dim iDot:iDot=InStrRev(s,".")
Dim iLen:iLen=Len(s)
If((iLen-iDot)>5)Or(iDot<1)Then
Extension=""
Else
Extension=Right(s,iLen-iDot)
End If
End Function
Function FileType(ByVal s):FileType=Extension(s):End Function
Function MimeType(ByVal sFT)
Dim sMIME
sMIME=LCase(FileType(sFT))
Select Case sMIME
Case "gif":sMIME="image/gif"
Case "jpg","jpe","jpeg":sMIME="image/jpeg"
Case "bmp":sMIME="image/bmp"
Case "png":sMIME="image/png"
Case "tif":sMIME="image/tiff"
End Select
MimeType=sMIME
End Function
Function ObfuscateCSS(ByVal s)
s=rxReplace(s,"/\*((?!\*/)[\S\s])*\*/","")
s=rxReplace(s,"\s*[\r\n]+\s*",vbCR&"")
s=rxReplace(s,"[\ \t]*([\{\}\(\;\:\,\!])[\ \t]*","$1")
s=rxReplace(s,"[\ \t]*([\)])","$1")
s=rxReplace(s,"[\ \t]+"," ")
s=rxReplace(s,"[\r\n]+","")
s=rxReplaceAll(s,"\s*\;\s*\}","}")
s=Replace(s,"and(","and (")
ObfuscateCSS=s
End Function
Function FWX__Geocode()
Dim ggeo
ggeo=Query.FQ("coords,address1,postcode,country")
If ggeo="" Then Exit Function
If ggeo="" Then
ggeo=Enclose(Query.FQ("address1"),";","")&Enclose(Query.FQ("city"),";","")&Enclose(Query.FQ("postcode"),";","")&Enclose(Query.FQ("country"),";","")&""
ggeo=rxReplace(rxReplace(Trim(ggeo),"^;+|;+$",""),";+",";")
End If
Dim gcrd,acrd,glat,glng
gcrd=rxSelect(ggeo,"[\d\.]+,[\d\.]+")
If gcrd="" Then
gcrd=GoogleCoords(ggeo)
End If
If gcrd="" Then
Exit Function
End If
acrd=Split(gcrd,",")
If UBound(acrd)=1 Then
glat=acrd(0)
glng=acrd(1)
End If
Query.Fata("coords")=gcrd
Query.Fata("coord")=gcrd
Query.Fata("lat")=glat
Query.Fata("lng")=glng
End Function
Function IIf(ByVal bln,ByVal returnTrue,ByVal returnFalse)
Err.Clear():On Error Resume Next
If IsSet(bln)Then
If IsTrue(bln)Then
IIf=CStr(returnTrue)
Else
IIf=CStr(returnFalse)
End If
Else
IIf=CStr(returnFalse)
End If
Err.Clear()
End Function
Function DoListOptions(ByRef ent,ByVal blnListRecurs,ByVal sCurrent,ByVal iRoot,ByVal sIndent)
On Error Resume Next
Err.Clear()
If Not IsObject(ent)Then Exit Function
With ent
Set rs=.RS()
If blnListRecurs Then rs.Filter="parent = "&iRoot
If Err.Description<>"" Or Err.Number<>0 Then Err.Clear():ent.Filter=""
If Not ExistsOrKill(rs)Then Exit Function
rs.Sort=TOT(Replace(ent.Sort,"ORDER BY",""),"[sort] ASC, [name] ASC")
If Err.Description<>"" Or Err.Number<>0 Then Err.Clear():ent.Sort=""
Do While Exists(rs)
output=output&"<option value="""&rs(.Key)&""""&DoSelect(rs(.Key)=sCurrent Or IsInList(rs(.Key),sCurrent))&">"
output=output&sIndent&" "
output=output&rs("name")
If Err.Description<>"" Or Err.Number<>0 Then Err.Clear():output=output&rs(ent.Key)
If Err.Description<>"" Or Err.Number<>0 Then Err.Clear():output=output&" * Error"
output=output&"</option>"
If blnListRecurs Then
sList=DoListOptions(ent,True,sCurrent,rs(.Key),sIndent&"...")
If IsSet(sList)Then
output=output&sList
output=output&"<option value=""""></option>"
End If
End If
rs.MoveNext
Loop
KillRS rs
End With
DoListOptions=output
End Function
Function ListOptions(ByRef ent)
ListOptions=DoListOptions(ent,True,Request(ent.Key),0,Empty)
End Function
Function ListOptionsFrom(ByRef ent,ByVal iRoot)
ListOptionsFrom=DoListOptions(ent,True,Request(ent.Key),iRoot,Empty)
End Function
Function ListOptionsWith(ByRef ent,ByRef objSource)
ListOptionsWith=DoListOptions(ent,True,TOT(objSource(ent.Key),Request(ent.Key)),0,Empty)
End Function
Function ListOptionsWithValue(ByRef ent,ByVal sCurrent)
ListOptionsWithValue=DoListOptions(ent,True,sCurrent,0,Empty)
End Function
Function ListOptionsNoRecur(ByRef ent)
ListOptionsNoRecur=DoListOptions(ent,False,Request(ent.Key),0,Empty)
End Function
Function ListOptionsNoRecurWith(ByRef ent,ByRef objSource)
ListOptionsNoRecurWith=DoListOptions(ent,False,TOT(objSource(ent.Key),Request(ent.Key)),0,Empty)
End Function
Function ListOptionsNoRecurWithValue(ByRef ent,sCurrent)
ListOptionsNoRecurWithValue=DoListOptions(ent,False,sCurrent,0,Empty)
End Function
Function EXEDBG(ByVal s)
Err.Clear():On Error Resume Next
DBG "##############################################"
DBG "##EXEDBG:"&Encode(s)
DBG "##############################################"
DBG s
DBG "##############################################"
DBG "##EXEDBG COMPLETE"
DBG "##############################################"
End Function
Function EXE(ByVal s)
If DBG.Show Then
DBG "##EXE:"&s
DBG ""
End If
If s<>"" Then
If rxTest(s,"\+")Then 'rxTest(s,"[^a-zA-Z0-9\_\(\)\""\'\.]")Then
DBG.Warning "EXE ERROR: "&s
Else
EXE=Eval(s)
End If
End If
End Function
Function EXE_Backward(ByVal s,ByVal f)
If f<>"" Then f="_"&f
Dim o
s=rxReplace(Trim(s),"[^\w\_]+","_")
Do While s<>"" 'And o=""
o=TOT(o,EXE(s&f))
If InStrRev(s,"_")>1 Then
s=Left(s,InStrRev(s,"_")-1)
Else
s=""
End If
Loop
EXE_Backward=o
End Function
Function EXE2(ByVal s)
Err.clear():On Error Resume Next
EXE2=Eval(s)
End Function
Function OBJ(ByVal s)
If IsObject(s)Then
Set OBJ=s
Else
s=TOE(s)
If s<>"" Then
Set OBJ=Eval(s)
End If
End If
End Function
Function OBJ2(ByVal s)
Err.clear():On Error Resume Next
If IsObject(s)Then
Set OBJ2=s
Else
s=TOE(s)
If s<>"" Then
Set OBJ2=Eval(s)
End If
If Err.Description<>"" Or Err.Number<>0 Then Set OBJ2=Nothing
End If
End Function
Const FWX_BlankGif="data:image/gif;base64,R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw=="
Function vsp(ByVal height)
vsp="<div style='clear:both;margin:0;padding:0;line-height:"& height&"px;'><img src='"&FWX_BlankGif&"' height='"& height&"' width='1' alt=''/></div>"
End Function
Function hsp(ByVal width)
hsp="<img src="""&FWX_BlankGif&""" width="""&width&""" height='1' alt='' style='clear:none;margin:0;padding:0;line-height:1px;' />"
End Function
Function BeginASP()
BeginASP=Replace("< marker %"," marker ","")
End Function
Function EndASP()
EndASP=Replace("% marker >"," marker ","")
End Function
Function Random(ByVal i)
Randomize()
Random=Int(Rnd()*VOZ(i))
End Function
Function DecodeCriteria(ByVal str)
str=Trim(Replace(Trim(Replace(Replace(Replace(str&""," "," "),"keywords",""),"filter results","")),","," "))
r.Pattern="(\'|\[|\"")([^\ \'\]|\""]+) ([^\ ]+)([\'|\]|\""| ])"
Do While r.Test(str)
str=r.Replace(str,"$1$2_$3$4")
Loop
r.Pattern="\'|\""|\[|\]"
str=r.Replace(str,"")
DecodeCriteria=str
End Function
Function ANDOR(ByVal s)
Select Case LCase(s)
Case "a"
ANDOR="AND"
Case "x"
ANDOR="XOR"
Case Else
ANDOR="OR"
End Select
End Function
Function MergeCriteria(ByVal a,ByVal z,ByVal b)
Dim o:o=""
z=TOT(z,"AND")
If a<>"" Then o=o&Trim(a)
If a<>"" And b<>"" Then o=o&" "&Trim(UCase(z&""))&" "
If b<>"" Then o=o&Trim(b)
MergeCriteria=o
End Function
Function EncloseCriteria(ByVal str)
If IsSet(str)Then EncloseCriteria="("&str&")"
End Function
Function WhereCriteria(ByVal str)
If IsSet(str)Then WhereCriteria=" WHERE ("&str&") "
End Function
Private Function IP2Num(ByVal ipA)
strO=ipA
pos1=InStr(strO,".")
intA=CInt(Left(strO,(pos1-1)))
strO2=Mid(strO,pos1+1,len(strO))
pos2=InStr(strO2,".")
intB=CInt(Left(strO2,(pos2-1)))
strO3=Mid(strO2,pos2+1,len(strO2))
pos3=InStr(strO3,".")
intC=CInt(Left(strO3,(pos3-1)))
intD=CInt(Mid(strO3,pos3+1,len(strO3)))
intConvert=(intA*(256*256*256))+(intB*(256*256))+(intC*256)+intD
IP2Num=Trim(intConvert)
End Function
Function OSTag(ByVal s)
Dim o,v
If o="" Then o=rxSelect(s,"\b(iPhone OS|Android|Nokia|SymbianOS)\s*([\d\.]+)?")
If o="" Then o=rxSelect(s,"\b(Windows NT|Mac OS X|Linux|Ubunto|Unix)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Windows|Macintosh)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(iPhone|BlackBerry)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Baidu)(spider)?")
If o="" Then o=rxSelect(s,"\b(DotBot)(spider)?")
If o="" Then o=IIf(rxTest(s,"\b(Alps VOYAGER DG\d+?)"),"Doogee",o)
If o<>"" Then
o=rxReplace(o,"NT 5$","XP")
If InStr(1,s,"NT 6.1")Then o=rxReplace(o,"NT 6$","7")
If InStr(1,s,"NT 6.2")Then o=rxReplace(o,"NT 6$","8")
If InStr(1,s,"NT 6.3")Then o=rxReplace(o,"NT 6$","8")
o=rxReplace(o,"NT 6$","Vista")
o=rxReplace(o,"NT 7$","8")
o=rxReplace(o,"NT 8$","9")
o=rxReplace(o,"[^\d\w\.](\d*)$","$1")
End If
OSTag=o
End Function
Function UATag(ByVal s)
Dim o,v,m
s=rxReplace(s,"\b(Mozilla).\d+(\.\d)?","")
s=rxReplace(s,"\s*\(compatible;\s*","")
s=rxReplace(s,"\s*\b(Windows NT \d.\d)\s*","")
s=Replace(s,".0 (compatible; ","")
s=Replace(s,".0+(compatible; ","")
If o="" And rxTest(s,"Trident/\d")Then
If rxTest(s,"Trident/5")Then o="MSIE 9"
If rxTest(s,"Trident/6")Then o="MSIE 10"
If rxTest(s,"Trident/7")Then o="MSIE 11"
If rxTest(s,"Trident/8")Then o="MSIE 12"
End If
If o="" Then o=rxSelect(s,"\b(iPhone|iPad)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Opera|Chrome)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Firefox|MSIE|Netscape)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Googlebot(-Image)?|MSNbot|Yahoo|Ask Jeeves)\b")
If o="" Then o=rxSelect(s,"\b(Nokia|NSeries).\d+")
If o="" Then o=rxSelect(s,"\b(Safari|Camino)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(AppleWebKit|Blackberry)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Apache)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Outlook|Thunderbird)\s*(\d+)?")
If o="" Then o=rxSelect(s,"\b(Baidu)(spider)?")
If o="" Then o=rxSelect(s,"\b(DotBot)(spider)?")
If o="" Then o=rxSelect(s,"\b(Testomato)?")
If o="" Then o=rxSelect(s,"\b(Uptimebot|UptimeRobot)?")
If o="" Then o=rxSelect(s,"\b(DG300)?")
If o="" Then o=rxSelect(s,"\b(MJ12)?")
If o<>"" Then
v=TOT(rxSelect(s,"Version/\d+"),o)
If rxTest(s,"iPad|iPhone")Then
ElseIf rxTest(s,"Safari|Firefox")Then
m=rxSelect(s,"\bMobile\b")
End If
o=m&rxReplace(o,"\D(\d+)$",rxSelect(v,"\d+$"))
End If
UATag=TOT(o,s)
End Function
Function EPDomain(ByVal s)
Dim o
If o="" Then o=rxSelect(s,"\b((gmail|googlemail|live|outlook|hotmail|yahoo)\.(com|co\.uk))\b")
EPDomain=o
End Function
Function EPTag(ByVal s)
Dim o,p
p=EPDomain(s)
If o="" And rxTest(p,"live|outlook|hotmail")Then o="hotmail"
If o="" And rxTest(p,"gmail|googlemail")Then o="gmail"
If o="" And rxTest(p,"yahoo")Then o="yahoo"
EPTag=o
End Function
Const FWX_DEFAULT_RECPP=50
Const FWX___PARSER_MAX_MERGE_DEPTH=25
Const FWX___PARSER_MAX_PARSE_DEPTH=25
Function ParseTemplateFile(ByVal o)
If StartsWith(o,"<!--EMPTY-->")Then o=""
If StartsWith(o,"<!--IGNORE-->")Then o=""
If StartsWith(o,"<!--MISSING-->")Then o=""
If StartsWith(o,"<!--DISABLED-->")Then o=""
If o<>"" Then o=parser.PreParse(o)
ParseTemplateFile=o
End Function
Function ReadTemplateFile(ByVal s)
Dim o
o=ReadFile(s)
If StartsWith(o,"<!--IGNORE-->")Then o=""
If StartsWith(o,"<!--MISSING-->")Then o=""
ReadTemplateFile=o
End Function
Function TemplateFile(ByVal s)
s=rxReplace(Trim(s),"\s+","/")
Dim o:o=""
Dim f
If o="" Then
f=("/$TMPL_FOLDER/html/"&s&".html")&""
o=ReadTemplateFile(f)
If o<>"" and IsTest Then
DBG "TemplateFile>1>"&f&"="&Len(o)&">"&Encode(o)
End If
End If
If o="" And TMPL_FOLDER<>SITE_FOLDER Then
f=("/$SITE_FOLDER/html/"&s&".html")&""
o=ReadTemplateFile(f)
If o<>"" and IsTest Then
DBG "TemplateFile>2>"&f&"="&Len(o)&">"&Encode(o)
End If
End If
If o="" Then
If FWX__TMPL__Folder<>"" Then
f=("/"&FWX__TMPL__Folder&"/"&s&".html")&""
o=ReadTemplateFile(f)
If o<>"" and IsTest Then
DBG "TemplateFile>3>"&f&"="&Len(o)&">"&Encode(o)
End If
End If
End If
If o="" Then
f=("/fwx/html/"&s&".html")&""
o=ReadTemplateFile(f)
If o<>"" and IsTest Then
DBG "TemplateFile>4>"&f&" ---> "&Len(o)&" bytes ---> "&Encode(o)
End If
End If
TemplateFile=o
End Function
Const rxTemplateField="\|[a-zA-Z0-9\[\]\/\_\:\,\.\-\@]+\|"
Function MergeTemplate(ByVal o,ByRef col)
If o="" Then
Else
o=rxReplace(o,"\#\|\*","#")
o=rxReplace(o,"\#\*\|","#")
Dim bTgs,i
bTgs=rxTest(o,rxTemplateField)
If Not bTgs Then
If IsObject(parser)Then
Dim bParsing,PrevText
bParsing=True
Do While bParsing
If o<>"" Then
PrevText=o
o=parser.PreParse(o)
bParsing=CBool(PrevText<>o)
Else
bParsing=False
End If
Loop
End If
bTgs=rxTest(o,rxTemplateField)
End If
i=0
Do While bTgs And Err.Number=0 And i<=FWX___PARSER_MAX_PARSE_DEPTH
i=i+1
If IsStaffPC And IsTest Then j("_STAT-MergeTemplate")=VOZ(j("_STAT-MergeTemplate"))+1
o=MergeTemplateStatic(o,col)
o=parser.PreParse(o)
Err.Clear()
bTgs=rxTest(o,rxTemplateField)
If i>FWX___PARSER_MAX_MERGE_DEPTH Then
End If
Loop
Err.Clear()
End If
MergeTemplate=o
End Function
Function MergeTemplateStatic(ByVal o,ByRef col)
Dim bTgs,i
i=0
s=rxReplace(o,"(<!--/?/?)?\{\!\{(/?/?-->)?((?!\}\!\})[\s\S])+(<!--/?/?)?\}\!\}(/?/?-->)?","")
bTgs=rxTest(s,rxTemplateField)
Do While bTgs And Err.Number=0 And i<=FWX___PARSER_MAX_MERGE_DEPTH
i=i+1
If IsStaffPC And IsTest Then j("_STAT-MergeTemplateStatic")=VOZ(j("_STAT-MergeTemplateStatic"))+1
o=MergeTemplateOnce(o,col)
s=rxReplace(o,"(<!--/?/?)?\{\!\{(/?/?-->)?((?!\}\!\})[\s\S])+(<!--/?/?)?\}\!\}(/?/?-->)?","")
bTgs=rxTest(s,rxTemplateField)
Loop
Err.Clear()
MergeTemplateStatic=o
End Function
Function MergeTemplateOnce(ByVal o,ByRef col)
Dim mf,mfs,rpc,fld
Set mfs=rxMatches(o,rxTemplateField)
For Each mf In mfs
If InStr(1,o,mf.Value)>0 Then
fld=(Mid(Trim(mf.Value),2,Len(mf.Value)-2))
rpc=MergeTemplateField(col,fld,TRUE)
o=Replace(o,mf.Value,rpc)
End If
Next
Err.Clear()
MergeTemplateOnce=o
End Function
Function MergeTemplateFieldContent(ByRef col,ByVal flds)
Dim rpc,atr,fld
fld=flds
If IsXml(col)Then
atr=rxSelect(fld,"\@.+$")
If atr<>"" Then
fld=Replace(fld,atr,"")
atr=Replace(atr,"@","")
rpc=XmlNodeAttrText(col,fld,atr)
Else
rpc=XmlNodeText(col,fld)
End If
Else
rpc=ColItemText(col,fld)
End If
If rpc="" Then
If InStr(1,flds,",")>0 Then
DBG "#MergeTemplateFieldContent# flds="&flds
For Each fld In Split(flds,",")
If rpc="" Then
DBG "#MergeTemplateFieldContent# fld="&fld
rpc=MergeTemplateFieldContent(col,fld)
DBG "#MergeTemplateFieldContent# fld="&fld&"=='"&Encode(rpc)&"'"
End If
Next
End If
End If
MergeTemplateFieldContent=rpc
End Function
Function MergeTemplateField(ByRef col,ByVal merge_string,ByVal exec_formats)
Dim fld,rpc,opt,par
fld=merge_string
opt=rxSelect(fld,"\:.+$")
If opt<>"" Then
fld=Mid(fld,1,Len(fld)-Len(opt))
opt=Mid(opt,2)
End If
par=rxSelect(fld,"\/.+$")
If par<>"" Then
fld=Mid(fld,1,Len(fld)-Len(par))
par=Mid(par,2)
End If
Select Case fld
Case "@isaddress":
rpc=IsAddress(col)
Case "@address":
rpc=DoAddress(col,TOT(opt,"text"))
Case "@account":
rpc=DoAccount(col,TOT(opt,"text"))
Case Else:
rpc=MergeTemplateFieldContent(col,fld)
If par<>"" Then
Dim dds
Set dds=Dictionary()
UpdateCollection dds,rpc
rpc=dds(par)
Set dds=Nothing
End If
If exec_formats Then
If opt<>"" Then
rpc=parser.Format(rpc,opt)
End If
End If
End Select
MergeTemplateField=rpc
End Function
Sub MergeAllStatic(ByRef dic,ByRef col)
Dim kkk
For Each kkk In dic
If IsObject(dic(kkk))Or DontMergeInBulk(kkk)Then
Else
dic(kkk)=MergeTemplateStatic(dic(kkk)&"",col)
End If
Next
End Sub
Sub MergeAllThatStartWith(ByRef dic,ByRef col,ByVal start)
Dim kkk
If start<>"" Then
For Each kkk In dic
If IsObject(dic(kkk))Or DontMergeInBulk(kkk)Then
Else
If dic(kkk)="" Then
ElseIf kkk="" Then
Else
If StartsWith(kkk,start)Then
dic(kkk)=MergeTemplate(dic(kkk)&"",col)
End If
End If
End If
Next
Else
MergeAll dic,col
End If
End Sub
Function DontMergeInBulk(ByVal kkk)
DontMergeInBulk=CBool(Not BulkMergeableField(kkk))
End Function
Function BulkMergeableField(ByVal kkk)
If False Then
ElseIf kkk="" Then
ElseIf rxTest(kkk,"^[\!\#\-\_\.\,\@]")Then
ElseIf rxTest(kkk,"^(e|my|sql)\b")Then
ElseIf rxTest(kkk,"access|credential")Then
ElseIf rxTest(kkk,"filter|criteria|search")Then
Else
BulkMergeableField=True
End If
End Function
Sub MergeAll(ByRef dic,ByRef col)
For Each kkk In dic
If IsObject(dic(kkk))Or DontMergeInBulk(kkk)Then
Else
If dic(kkk)="" Then
ElseIf kkk="" Then
Else
dic(kkk)=MergeTemplate(dic(kkk)&"",col)
End If
End If
Next
End Sub
Sub SelfMerge(ByRef dic)
For Each kkk In dic
If IsObject(dic(kkk))Or DontMergeInBulk(kkk)Then
Else
dic(kkk)=MergeTemplateStatic(dic(kkk)&"",dic)
End If
Next
End Sub
Function MergeRS(ByRef rs,ByVal sT)
Dim o,i,c
Set o=New QuickString
i=0
If Exists(rs)Then
Do While Exists(rs)
Set c=RS2Dic(rs)
i=i+1
c("i")=i
o.Add MergeTemplate(sT,c)
rs.MoveNext
Loop
End If
MergeRS=o
End Function
Sub MergeDic(ByRef dic,ByRef lclSource)
For Each kkk In dic
If IsObject(dic(kkk))Or DontMergeInBulk(kkk)Then
Else
dic(kkk)=MergeTemplateStatic(dic(kkk)&"",lclSource)
End If
Next
End Sub
Sub MergeDicLive(ByRef dic,ByRef lclSource)
For Each kkk In dic
If IsObject(dic(kkk))Or DontMergeInBulk(kkk)Then
Else
dic(kkk)=MergeTemplate(dic(kkk)&"",lclSource)
End If
Next
End Sub
Function MT(ByRef col,ByVal str)
If IsObject(col)Then
MT=MergeTemplate(str,col)
Else
MT=MergeTemplate(col,str)
End If
End Function
Function MJ(ByVal str)
MJ=MT(j,str)
End Function
Function Transfer(ByVal s)
Dim b,target_extension
s=rxReplace(s,"[\\\/]+","/")
On Error Resume Next
target_extension=LCase(Replace(rxSelect(s,"\.\w{2,4}$"),".",""))
Select Case target_extension
Case "js","json","css","htm","html","txt","xml":
Server.Transfer s
Case "png","jpg","gif","jpe","bmp","jpeg","tiff":
Response.ContentType="image/"&rxReplace(target_extension,"jpg?e?","jpeg")
Err.Clear():On Error Resume Next
Dim d:d=ReadBinary(s)
If Len(d)>0 And Err.Number=0 Then
RC
Response.BinaryWrite d
RE
End If
FWX_ResponseStatus "301 Moved Permanently"
FWX_AddHeader "Location",s
Case "asp":
Server.Transfer s
Case Else
PermanentRedirect s
End Select
If Err.Number&""<>"-2147467259" Then
Response.End()
End If
Transfer=CBool(Err.Number=0)
Err.Clear()
End Function
Function Miles2Km():Miles2Km=(1.60934):End Function
Function Km2Miles():Km2Miles=(1/1.60934):End Function
Function Metres2Inches():Metres2Inches=(39.37008):End Function
Function Metres2Feet():Metres2Feet=(Metres2Inches()*Inches2feet()):End Function
Function Metres2Yards():Metres2Yards=(Metres2Inches()*Inches2Yards()):End Function
Function Metres2Miles():Metres2Miles=(Metres2Inches()*Inches2Miles()):End Function
Function Inches2Feet():Inches2Feet=(1/12):End Function
Function Inches2Metres():Inches2Metres=(0.254):End Function
Function Inches2Miles():Inches2Miles=(1/(12*5280)):End Function
Function Inches2Yards():Inches2Yards=(1/(12*3)):End Function
Function Yards2Inches():Yards2Inches=(12*3):End Function
Function Yards2Metres():Yards2Metres=(Yards2Inches()*Inches2Metres()):End Function
Function Feet2Inches():Feet2Inches=(12):End Function
Function Feet2Miles():Feet2Miles=(1/5280):End Function
Const GZipKey="\x1f\x8b\x08\x00\x00\x00\x00\x00"
Const GZipScript="/fwx/php/gzip.php"
Function GZipPath()
GZipPath=MyDomainCache&"gzip/"
End Function
Dim GZipAcceptedCache
Function GZipAccepted()
Exit Function
If GZipAcceptedCache="" Then
GZipAcceptedCache=" "
If 1=0 Then
ElseIf IsMobile Then
ElseIf IsIE And(IEVersion<=6 Or IEVersion=8)Then
Else
If USE_GZIP Or IsTrue(FWX.Setting("gzip"))Then
Dim s:s=Request.ServerVariables("HTTP_ACCEPT_ENCODING")
If InStr(1,s,"x-gzip")>0 Then GZipAcceptedCache="x-gzip"
If InStr(1,s,"gzip")>0 Then GZipAcceptedCache="gzip"
End If
End If
End If
GZipAccepted=Trim(GZipAcceptedCache)
End Function
Function GZipFile(ByVal s)
If GZipAccepted="" Then Exit Function
Dim dat
Dim out,key
key=DoEncrypt(MD5(s),"1234567890")
out=GZipPath()&key&".gz"
BuildPath GZipPath()
ed=FWX.Setting("edition")
ca=FWX.Setting("GZIP.Cache")
If Left(ca,Len(ed))=ed Then
If InStr(1,ca,key&",")>0 Then
dat=ReadBinary(out)
Else
ca=ca&key&","
End If
Else
ca=ed&","&key&","
End If
FWX.Setting("GZIP.Cache")=ca
If Len(dat)>0 Then
Else
Dim url,str
url="http://"&SERVER_NAME&GZipScript&"?e="&GZipAccepted&"&"
url=url&"u="&s&"&f="&out&"&"
str=GetResponse(url)
If 1=0 Then
ElseIf InStr(1,str,"PHP Warning")>0 Then
ElseIf InStr(1,str,"PHP Error")>0 Then
ElseIf InStr(1,str,"Parse error")>0 Then
ElseIf UCase(Left(str,2))="OK" Then
dat=ReadBinary(out)
End If
End If
GZipFile=dat
End Function
Function GZipContent(ByVal sContent)
If GZipAccepted="" Then Exit Function
sContent=Trim(sContent)
If sContent="" Then Exit Function
If Len(sContent)<=1024 Then DIE sContent
Dim GZk,s,out,dat,gZipT:gZipT=Timer()
GZk=MD5(sContent&"")
s=GZipPath()&GZk&".txt"
DeleteFile s
WriteToFile s,sContent
dat=GZipFile(s)
DeleteFile s
If Len(dat)>0 Then
Response.Clear()
FWX_AddHeader "Content-Encoding",GZipAccepted
Response.BinaryWrite dat
Response.Flush()
Response.End()
Else
End If
End Function
Function GZipUrl(ByVal sUrl)
If GZipAccepted="" Then Exit Function
sUrl=FullyQualifiedUrl(Trim(sUrl))
If sUrl="" Then Exit Function
Dim dat,gZipT:gZipT=Timer()
dat=GZipFile(sUrl)
If Len(dat)>0 Then
Response.Clear()
FWX_AddHeader "Content-Encoding",GZipAccepted
Response.BinaryWrite dat
Response.Flush()
Response.End()
Else
End If
End Function
Const ForReading=1
Const ForWriting=2
Const ForAppending=8
Const TristateUseDefault=-2
Const TristateTrue=-1
Const TristateFalse=0
Const adTypeBinary=1
Const adTypeText=2
Function FixPath(ByVal s)
s=rxReplace(s&"","[\\\/]+","/")
FixPath=s
End Function
Function FileName(ByVal s)
Dim c:c="[^\w\-\_]+"
s=rxReplace(s,"'s\b","s")
s=rxReplace(s,c&"$","")
s=rxReplace(s,"^"&c,"")
s=rxReplace(s,c,"-")
FileName=s
End Function
Function SafeFileName(ByVal s)
Dim b:b="[\""\'\/\\\:\|]"
s=rxReplace(s,b,"-")
SafeFileName=s
End Function
Function ArchivePath(ByVal s)
ArchivePath=ArchivePath10(s)
End Function
Function ArchivePath10(ByVal s)
Dim o:o=""'/"
Dim i:i=0
For i=1 To Len(s)-1
o=o&Left(s,i)&"/"
Next
o=o&s
ArchivePath10=o
End Function
Function ArchivePath100(ByVal s)
Dim o:o=""'/"
Dim i:i=0
Do While Len(s)>2
o=o&Left(s,2)&"/"
s=Mid(s,3)
Loop
o=o&s
ArchivePath100=o
End Function
Function ArchivePath1000(ByVal s)
Dim o:o=""'/"
Dim i:i=0
Do While Len(s)>3
o=o&Left(s,3)&"/"
s=Mid(s,4)
Loop
o=o&s
ArchivePath1000=o
End Function
Function ArchivePath10000(ByVal s)
Dim o:o=""'/"
Dim i:i=0
Do While Len(s)>4
o=o&Left(s,4)&"/"
s=Mid(s,5)
Loop
o=o&s
ArchivePath10000=o
End Function
Function ArchivePath100000(ByVal s)
Dim o:o=""'/"
Dim i:i=0
Do While Len(s)>5
o=o&Left(s,5)&"/"
s=Mid(s,6)
Loop
o=o&s
ArchivePath100000=o
End Function
Function ArchivePath1(ByVal s):ArchivePath1=ArchivePath10(s):End Function
Function ArchivePath2(ByVal s):ArchivePath2=ArchivePath100(s):End Function
Function ArchivePath3(ByVal s):ArchivePath3=ArchivePath1000(s):End Function
Function ArchivePath4(ByVal s):ArchivePath4=ArchivePath10000(s):End Function
Function ArchivePath5(ByVal s):ArchivePath5=ArchivePath100000(s):End Function
Function SafePath(ByVal s)
SafePath=Encrypt(Reverse(ArchivePath(Left(RemoveSymbols(s),8))))
End Function
Function FullyQualifiedUrl(ByVal s)
s=Trim(s&"")
If InStr(1,s,"://")>0 Then
Else
If Left(s,1)<>"/" Then
s=PATH&s
End If
s=MyProtocol&"://"&SERVER_NAME&s&""
End If
FullyQualifiedUrl=s
End Function
Function VirtualPath(ByVal s)
s=PhysicalPath(s)
If rxTest(s,"\\fwx\\")Then
s=rxSelect(s,"^.+(\\fw\\.+)$","$1")
Else
s=Mid(s,Len(Server.MapPath("/")&"\"))
End If
s=Replace(s,"\","/")
VirtualPath=s
End Function
Function FolderPath(ByVal s)
s=FixPath(s)
s=Left(s,InStrRev(s,"/"))
FolderPath=s
End Function
Function ServerMapPath(ByVal s)
Err.Clear():On Error Resume Next
s=FixPath(s)
If s<>"" Then
s=Replace(Server.URLEncode(s),"%2F","/")
s=Server.MapPath(s)
s=URLDecode(s)
If Err.Number<>0 Then
DBG.FatalError "Server.MapPath Error:"&Err.Description&" / PATH="&s
End If
End If
ServerMapPath=s
End Function
Function DecodePath(ByVal s)
DecodePath=TranslatedPath(s)
End Function
Function TranslatedPath(ByVal s)
If InStr(1,s,"$ROOT")>0 Then s=Replace(s,"$ROOT",ROOT)
If InStr(1,s,"$BELOWROOT")>0 Then s=Replace(s,"$BELOWROOT",BELOWROOT)
If InStr(1,s,"$SITE_FOLDER")>0 Then s=Replace(s,"$SITE_FOLDER",SITE_FOLDER)
If InStr(1,s,"$TMPL_FOLDER")>0 Then s=Replace(s,"$TMPL_FOLDER",TMPL_FOLDER)
If InStr(1,s,"$DOMAINCACHE")>0 Then s=Replace(s,"$DOMAINCACHE",MyDomainCache)
If InStr(1,s,"$FE__TMPL_FOLDER")>0 Then s=Replace(s,"$FE__TMPL_FOLDER",FE__TMPL_FOLDER)
If InStr(1,s,"$BO__TMPL_FOLDER")>0 Then s=Replace(s,"$BO__TMPL_FOLDER",BO__TMPL_FOLDER)
s=rxReplace(s&"","[\\\/]+","/")
TranslatedPath=s
End Function
Function PhysicalPath(ByVal s)
s=TranslatedPath(s)
If InStr(1,s,"/")>0 Then
If Mid(s,2,1)<>":" Then
s=ServerMapPath(s)
End If
End If
PhysicalPath=s
End Function
Function FSOFolder(ByVal sFD)
GetFolder FSOFolder,sFD
End Function
Function GetFolder(ByRef oTargetObject,ByVal sFD)
Err.Clear():On Error Resume Next
Dim oFSO:Set oFSO=FSObject()
Set oTargetObject=oFSO.GetFolder(PhysicalPath(sFD))
Set oFSO=Nothing
GetFolder=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function GetFile(ByRef oTargetObject,ByVal sFL)
Err.Clear():On Error Resume Next
j("_STAT-FSO.GetFile")=VOZ(j("_STAT-FSO.GetFile"))+1
j("_STAT-FSO.GetFile-List")=j("_STAT-FSO.GetFile-List")&"|||"&sFL
Dim oFSO:Set oFSO=FSObject()
Set oTargetObject=oFSO.GetFile(PhysicalPath(sFL))
Set oFSO=Nothing
GetFile=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function GetParentFolder(ByVal sFD)
Dim oFSO:Set oFSO=FSObject()
GetParentFolder=oFSO.GetParentFolderName(sFD)
Set oFSO=Nothing
End Function
Function FSOFiles(ByVal sFD)
Dim d,f,x
Set d=FSOFolder(sFD)
Set x=Dictionary()
For Each f In d.files
x(f.Name)=f.Size
Next
Set f=Nothing
Set d=Nothing
Set FSOFiles=x
End Function
Function GetFiles(ByRef oRS,ByVal sFD)
Dim oFolder
If IsObject(sFD)Then
Set oFolder=sFD
Else
GetFolder oFolder,sFD
End If
Set oRS=Server.CreateObject("ADODB.Recordset")
If IsObject(oFolder)Then
oRS.Fields.Append "path",adVarChar,250
oRS.Fields.Append "folder",adVarChar,250
oRS.Fields.Append "file",adVarChar,250
oRS.Fields.Append "name",adVarChar,250
oRS.Fields.Append "ext",adVarChar,20
oRS.Fields.Append "nam",adVarChar,250
oRS.Fields.Append "size",adInteger
oRS.Fields.Append "size_caption",adVarChar,50
oRS.Fields.Append "date",adDate
oRS.Fields.Append "modified",adDate
oRS.Fields.Append "type",adVarChar,255
oRS.Fields.Append "i",adVarChar,4
oRS.Open
Dim i:i=1
For Each oFile In oFolder.Files
oRS.AddNew
oRS.Fields("i")=i
i=i+1
oRS.Fields("path")=sFD
oRS.Fields("folder")=sFD
oRS.Fields("file")=oFile.Name
oRS.Fields("name")=oFile.Name
oRS.Fields("ext")=""
oRS.Fields("nam")=oFile.Name
If InStr(1,oFile.Name,".")>0 And InStr(1,oFile.Name,".")<Len(oFile.Name)Then
oRS.Fields("ext")=Mid(oFile.Name,InStrRev(oFile.Name,".")+1)
oRS.Fields("nam")=Mid(oFile.Name,1,Len(oFile.Name)-(Len(oRS.Fields("ext"))+1))
End If
oRS.Fields("size")=oFile.Size
If oFile.Size<1024 Then
oRS.Fields("size_caption")=oFile.Size&" bytes"
Else
oRS.Fields("size_caption")=FormatNumber(oFile.Size/1024,1)
End If
oRS.Fields("date")=oFile.DateCreated
oRS.Fields("modified")=oFile.DateLastModified
oRS.Fields("type")=oFile.Type
oRS.MoveLast
Next
If Exists(oRS)Then oRS.MoveFirst
End If
End Function
Function GetFilesInFolder(ByVal sFD)
GetFiles GetFilesInFolder,sFD
End Function
Function GetFilesOfType(ByVal sFD,ByVal sExt)
Dim rA:Set rA=GetFilesInFolder(sFD)
If Exists(rA)Then
If sExt<>"" Then
Do While Exists(rA)
If Not rxTest(rA("name"),"\.("&sExt&")$")Then
rA.Delete
End If
rA.MoveNext
Loop
If rA.RecordCount>0 Then rA.MoveFirst()
End If
End If
Set GetFilesOfType=rA
End Function
Function FileExists(ByVal sFL)
Dim f,b
f=sFL
Dim path
If InStr(1,f,"/")>0 Then path=Left(f,InStrRev(f,"/")-0)
If InStr(1,f,"\")>0 Then path=Left(f,InStrRev(f,"\")-0)
f=Replace(TOE(f),"%20"," ")
If path<>"" Then
If FolderExists(path)Then
Else
FileExists=False
Exit Function
End If
End If
Dim f_check
f_check=PhysicalPath(f)
Dim oFSO:Set oFSO=FSObject()
b=oFSO.FileExists(f_check)
Set oFSO=Nothing
j("_STAT-FSO.FileExists")=VOZ(j("_STAT-FSO.FileExists"))+1
FileExists=b
End Function
Function FolderExistsRAW(ByVal folder)
Dim f_check:f_check=PhysicalPath(folder)
Dim obj_fso:Set obj_fso=FSObject()
Dim f_exist:f_exist=obj_fso.FolderExists(f_check)
Set obj_fso=Nothing
j("_STAT-FSO.FolderExists")=VOZ(j("_STAT-FSO.FolderExists"))+1
j("_STAT-FSO.FolderExists-List")=j("_STAT-FSO.FolderExists-List")&"|||"&folder
FolderExistsRAW=f_exist
End Function
Dim FOLDERS___Y:FOLDERS___Y=""
Dim FOLDERS___N:FOLDERS___N=""
Function FolderExists(ByVal sFD)
Dim f,b
b=""
f=sFD
f=PhysicalPath(f)
Dim dirs,node,foot,path,fath
foot=TrueRoot(sFD)
dirs=Mid(f,Len(foot)+2)
fath="\"&dirs&"\"
If dirs="" Then
b=True
Else
If Left(dirs,1)="\" Then dirs=Mid(dirs,2)
If b="" Then
path=fath
If InStr(1,";"&FOLDERS___N&";",";"&path&";")>0 Then
b=False
End If
End If
If b="" Then
path=fath
If InStr(1,";"&FOLDERS___Y&";",";"&path&";")>0 Then
b=True
End If
End If
If b="" Then
path="\"
For Each node In Split(dirs,"\")
path=path&node&"\"
If InStr(1,";"&FOLDERS___N&";",";"&path&";")>0 Then
b=False
End If
Next
End If
If b="" Then
path=fath
If FolderExistsRAW(path)Then
path="\"
For Each node In Split(dirs,"\")
path=path&node&"\"
FOLDERS___Y=FOLDERS___Y&";"&path
Next
b=True
Else
FOLDERS___N=FOLDERS___N&";"&path
path="\"
For Each node In Split(dirs,"\")
path=path&node&"\"
If path=fath Then
b=False
Else
If b<>"" Then
FOLDERS___N=FOLDERS___N&";"&path
ElseIf 1=0 Or rxTest(Replace(path,"\","/"),"^/fwx(/(html|inc|js)/fe)?/$")Or rxTest(Replace(path,"\","/"),"^/fwx(/(html|inc|js)/bo)?/$")Then
ElseIf 1=0 Or path="\$$TEMP\" Or rxTest(Replace(path,"\","/"),"^/fwx(/(html|@CODE(/(fe|bo))?|js(/(plugins|connectors))?|html(/(shop(/output)?|quot(/output)?|staff|mail|dump))?|inc))?/$")Or rxTest(Replace(path,"\","/"),"^/fwx(/fe(/html(/(page|product|@|@/(core|ACTIONS|PAGES|ERRORS|-RESPONSE|-TMPL|bo|bo/bar)|client|cart|do|frontpage))?)?)?/$")Then
Else
If 1=0 Then
ElseIf InStr(1,";"&FOLDERS___N&";",";"&path&";")>0 Then
b=False
ElseIf InStr(1,";"&FOLDERS___Y&";",";"&path&";")>0 Then
Else
If FolderExistsRAW(path)Then
FOLDERS___Y=FOLDERS___Y&";"&path
Else
FOLDERS___N=FOLDERS___N&";"&path
b=False
End If
End If
End If
End If
Next
End If
If b="" Then b=True
End If
End If
FolderExists=b
End Function
Function TrueRoot(ByVal s)
Dim p,r
s=FixPath(s)
p=ServerMapPath(s)
If Right(p,1)="\" Then p=Mid(p,1,Len(p)-1)
s=Replace(s,"/","\")
If Right(s,1)="\" Then s=Mid(s,1,Len(s)-1)
r=Len(p)-Len(s)
If r>0 Then
s=Left(p,Len(p)-Len(s))
Else:RTXT:DIE "*************************************** ERROR:"&s:End If
TrueRoot=s
End Function
Function CreateFolder(ByVal sFD)
Err.Clear():On Error Resume Next
Dim oFSO:Set oFSO=FSObject()
oFSO.CreateFolder(PhysicalPath(sFD))
Set oFSO=Nothing
CreateFolder=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function BuildPath(ByVal sFD)
Err.Clear():On Error Resume Next
Dim oFSO:Set oFSO=FSObject()
Dim sPath,sFE
sPath="/"
For Each sFD In Split(sFD,"/")
If Trim(sFD)<>"" Then
sPath=sPath&Trim(sFD)&"/"
sFD=PhysicalPath(sPath)
sFE=oFSO.FolderExists(sFD)
If Not sFE Then oFSO.CreateFolder sFD
End If
Next
Set oFSO=Nothing
BuildPath=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function NewFile(ByRef oTargetObject,ByVal sFL)
Err.Clear():On Error Resume Next
Dim oFSO:Set oFSO=FSObject()
sFL=PhysicalPath(sFL)
BuildPath sFL
Set oTargetObject=oFSO.CreateTextFile(sFL,True)
Set oFSO=Nothing
If Err.Number<>0 Then DBG.Warning "FSO NewFile Error "&sFL
End Function
Function CreateFile(ByVal sFL)
NewFile tmp_file,sFL
Set tmp_file=Nothing
CreateFile=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function LastModified(ByVal sFL)
If FileExists(sFL)Then
sFL=PhysicalPath(sFL)
Dim oFL
If GetFile(oFL,sFL)Then
LastModified=oFL.DateLastModified
End If
Set oFL=Nothing
End If
End Function
Function FileSize(ByVal sFL)
If FileExists(sFL)Then
sFL=PhysicalPath(sFL)
Dim oFL
If GetFile(oFL,sFL)Then
FileSize=oFL.Size
End If
Set oFL=Nothing
End If
End Function
Function GetFileSize(ByVal sFL)
GetFileSize=FileSize(sFL)
End Function
Function makeArray(n)
Dim s
s=Space(n)
makeArray=Split(s," ")
End Function
Function GetFileForReading(ByRef oTargetObject,ByVal sFL)
Err.Clear():On Error Resume Next
j("_STAT-FSO.GetFile")=VOZ(j("_STAT-FSO.GetFile"))+1
j("_STAT-FSO.GetFileForReading")=VOZ(j("_STAT-FSO.GetFileForReading"))+1
j("_STAT-FSO.GetFileForReading-List")=j("_STAT-FSO.GetFileForReading-List")&"|||"&sFL
Dim path:path=PhysicalPath(sFL)
Dim oFSO:Set oFSO=FSObject()
Set oTargetObject=oFSO.OpenTextFile(path,ForReading,False)
Set oFSO=Nothing
GetFileForReading=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function GetFileForWriting(ByRef oTargetObject,ByVal sFL)
Err.Clear():On Error Resume Next
j("_STAT-FSO.GetFile")=VOZ(j("_STAT-FSO.GetFile"))+1
j("_STAT-FSO.GetFileForWriting")=VOZ(j("_STAT-FSO.GetFileForWriting"))+1
j("_STAT-FSO.GetFileForWriting-List")=j("_STAT-FSO.GetFileForWriting-List")&"|||"&sFL
Dim path:path=PhysicalPath(sFL)
Dim oFSO:Set oFSO=FSObject()
Set oTargetObject=oFSO.OpenTextFile(path,ForWriting,True)
Set oFSO=Nothing
GetFileForWriting=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function ReadFileFast(ByVal sFL)
ReadFileFast=ReadFileWithoutValidation(sFL)
End Function
Function ReadFileWithoutValidation(ByVal sFL)
Dim FSO_ReadFile_START
FSO_ReadFile_START=Timer()
sFL=TranslatedPath(sFL)
sFL=PhysicalPath(sFL)
Dim s,i,oFL,oST
If GetFileForReading(oST,sFL)Then
Set s=New QuickString
s.Divider=vbCRLf
Do While Not oST.AtEndOfStream
s.Add oST.ReadLine
Loop
oST.Close
Set oST=Nothing
j("_STAT-FSO.ReadFile")=VOZ(j("_STAT-FSO.ReadFile"))+1
j("_STAT-FSO.ReadFile-List")=j("_STAT-FSO.ReadFile-List")&"|||"&VirtualPath(sFL)
End If
ReadFileWithoutValidation=s
End Function
Function ReadFile(ByVal sFL)
sFL=TranslatedPath(sFL)
Dim sPF:sPF=rxReplace(sFL,"/[^/]+$","")
Dim s:s=""
If FolderExists(sPF)Then
s=ReadFileWithoutValidation(sFL)
Else
End If
DBG "!!! FSO.ReadFile: "&TranslatedPath(sFL)&"="&Len(s)&"b"
ReadFile=s
End Function
Function TextStream(ByRef oFL)
Dim oTargetObject
Set oTargetObject=Nothing
OpenTextStream oTargetObject,oFL
Set TextStream=oTargetObject
End Function
Function OpenTextStream(ByRef oTargetObject,ByRef oFL)
On Error Resume Next
Set oTargetObject=oFL.OpenAsTextStream(ForWriting,TristateUseDefault)
If Err.Number<>0 Then Err.Clear
End Function
Function WriteToFile(ByVal sFL,ByVal sDATA)
Dim sFD,oST,oFL
sFL=TranslatedPath(sFL)
If sFL="" Then
Exit Function
End If
sFD=Left(sFL,InStrRev(sFL,"/"))
If Not FolderExists(sFD)Then BuildPath sFD
If GetFileForWriting(oST,sFL)Then
oST.Write sDATA
oST.Close
Set oST=Nothing
WriteToFile=CBool(Not(Err.number>0))
Else
WriteToFile=False
End If
If Err.Description<>"" Or Err.Number<>0 Then
Err.Clear()
End If
End Function
Function ReadBinary(ByVal sFL)
sFL=PhysicalPath(sFL)
Dim oST
Set oST=Server.CreateObject("ADODB.Stream")
oST.Type=1
oST.Open
oST.LoadFromFile sFL
ReadBinary=oST.Read
oST.Close
Set oST=Nothing
End Function
Function WriteBinary(ByVal sFL,ByVal bData)
Dim sFD:sFD=Left(sFL,InStrRev(sFL,"/"))
If Not FolderExists(sFD)Then BuildPath sFD
sFL=PhysicalPath(sFL)
DeleteFile sFL
Dim oST
Set oST=Server.CreateObject("ADODB.Stream")
oST.type=1
oST.mode=3
oST.open
oST.write bData
oST.Position=0
oST.SaveToFile sFL
oST.Close
Set oST=Nothing
WriteBinary=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function MoveFolder(ByVal sFD,ByVal sDest)
Err.Clear():On Error Resume Next
Dim oFSO:Set oFSO=FSObject()
oFSO.MoveFolder PhysicalPath(sFD),PhysicalPath(sDest)&"\"
Set oFSO=Nothing
MoveFolder=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function MoveFile(ByVal sFL,ByVal sDest)
Err.Clear():On Error Resume Next
Dim oFSO:Set oFSO=FSObject()
oFSO.MoveFile PhysicalPath(sFL),PhysicalPath(sDest)&"\"
Set oFSO=Nothing
MoveFile=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function DoCopyFile(ByVal Source,ByVal sDest,ByVal NewName,ByVal DeleteSource)
Dim oFL
If GetFile(oFL,Source)Then
Dim cp,cn
If NewName<>"" Then
cn=NewName
cp=sDest
Else
cn=rxSelect(sDest,"[^/]*$")
cp=rxReplace(sDest,"[^/]*$","")
End If
BuildPath cp
oFL.Copy PhysicalPath(cp)&"\"&cn
End If
Set oFL=Nothing
If DeleteSource And(Not(Err.number>0))Then DeleteFile Source
DoCopyFile=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function CopyFileRename(ByVal Source,ByVal sDest,ByVal NewName)
CopyFileRename=DoCopyFile(Source,sDest,NewName,False)
End Function
Function CopyFile(ByVal Source,ByVal sDest)
CopyFile=DoCopyFile(Source,sDest,Empty,False)
End Function
Function RenameFileLeaveSource(ByVal Source,ByVal NewName)
Location=Left(Source,InStrRev(Source,"/"))
RenameFileLeaveSource=DoCopyFile(Source,Location,NewName,False)
End Function
Function RenameFile(ByVal Source,ByVal NewName)
Location=Left(Source,InStrRev(Source,"/"))
If Source<>(Location&NewName)Then
RenameFile=DoCopyFile(Source,Location,NewName,True)
End If
End Function
Function CopyFolderContents(ByVal Source,ByVal sDest)
Err.Clear():On Error Resume Next
Dim oFD
If GetFolder(oFD,Source)Then
For Each File In oFD.Files
Errors=Errors And CopyFile(Source&File.Name,sDest)
Next
End If
Set oFD=Nothing
CopyFolderContents=CBool((Not(Err.number>0))And(Not Errors))
Err.Clear()
End Function
Function DoCopyFolder(ByVal Source,ByVal sDest,ByVal NewName,ByVal DeleteSource)
Err.Clear():On Error Resume Next
Dim oFD
If GetFolder(oFD,Source)Then
oFSO.Copy PhysicalPath(sDest)&"\"&NewName
End If
Set oFD=Nothing
If DeleteSource And(Not(Err.number>0))Then DeleteFolder Source
DoCopyFolder=CBool(Not(Err.number>0)):Err.Clear()
End Function
Function CopyFolderRename(ByVal Source,ByVal sDest,ByVal NewName)
CopyFolderRename=DoCopyFolder(Source,sDest,NewName,False)
End Function
Function CopyFolder(ByVal Source,ByVal sDest)
CopyFolder=DoCopyFolder(Source,sDest,Empty,False)
End Function
Function RenameFolderLeaveSource(ByVal Source,ByVal NewName)
Location=Left(Source,InStrRev(Source,"/"))
RenameFolderLeaveSource=DoCopyFolder(Source,Location,NewName,False)
End Function
Function RenameFolder(ByVal Source,ByVal NewName)
Location=Left(Source,InStrRev(Source,"/"))
If Source<>Location&NewName Then
RenameFolder=DoCopyFolder(Source,Location,NewName,True)
Else
RenameFolder=True
End If
End Function
Function DeleteFolder(ByVal sFD)
If InStr(1,sFD,"/")>0 Then sFD=PhysicalPath(sFD)
Dim bDel
Dim oFSO:Set oFSO=FSObject()
If oFSO.FolderExists(sFD)Then
bDel=oFSO.DeleteFolder(sFD,True)
End If
Set oFSO=Nothing
DeleteFolder=bDel
End Function
Function DeleteFile(ByVal sFL)
Err.Clear():On Error Resume Next
If InStr(1,sFL,"/")>0 Then sFL=PhysicalPath(sFL)
Dim oFSO:Set oFSO=FSObject()
oFSO.DeleteFile(sFL)
Set oFSO=Nothing
Dim bResult:bResult=CBool(Not(Err.number>0)):Err.Clear()
Err.Clear()
DeleteFile=bResult
End Function
Dim LastHTTPRequest
Function HTTPRequestURL(ByVal sURL)
If sURL<>"" Then
sURL=FullyQualifiedUrl(sURL)
If rxTest(sURL,"^https?\:\/\/"&rxEscape(SERVER_NAME)&"\/")Then
If rxTest(sURL,"\?$")Or rxTest(sURL,"\&$")Then
Else:sURL=sURL&IIf(rxTest(sURL,"\?"),"&","?"):End If
sURL=rxReplace(sURL,"\b(ajax|v|auth|MyState|HTTPReq)\=[^&]+\&","")
If InStr(1,sURL,"/fwx/php/")>0 Then
Else
If InStr(1,sURL,"/x/")>0 Then
End If
End If
End If
sURL=rxReplace(sURL,"[\?\&]+$","")
End If
HTTPRequestURL=sURL
End Function
Const FWX_HTTPVerb_Async=True
Function HTTPVerb(ByVal sVERB,ByVal sURL)
Set HTTPVerb=OpenHTTPVerb(sVERB,sURL,FWX_HTTPVerb_Async,Null,Null)
End Function
Function SyncHTTPVerb(ByVal sVERB,ByVal sURL)
Set SyncHTTPVerb=OpenHTTPVerb(sVERB,sURL,False,Null,Null)
End Function
Function ASyncHTTPVerb(ByVal sVERB,ByVal sURL)
Set ASyncHTTPVerb=OpenHTTPVerb(sVERB,sURL,True,Null,Null)
End Function
Function OpenHTTPVerb(ByVal sVERB,ByVal sURL,ByVal ASyncMode,ByVal sUSR,ByVal sPWD)
Call CheckClientIsConnected()
sURL=HTTPRequestURL(sURL)
LastHTTPRequest=sURL
If sURL="" Then
DBG.FatalError "HTTP"&sVERB&" aborted (empty sURL)"
Else
If IsTest Then j("___last_http_"&sVERB)=sURL
j("_STAT-HTTP"&sVERB)=VOZ(j("_STAT-HTTP"&sVERB))+1
Set OpenHTTPVerb=Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
OpenHTTPVerb.Open sVERB,sURL,ASyncMode,sUSR,sPWD
OpenHTTPVerb.SetRequestHeader "Content-Type","application/x-www-form-urlencoded; charset=utf-8"
OpenHTTPVerb.SetRequestHeader "Accept-Charset","utf-8"
OpenHTTPVerb.SetRequestHeader "X-Requested-With","HTTPRequest"
End If
End Function
Function ReadHTTPVerb(ByVal sVERB,ByVal sURL)
Dim oTransport
Set oTransport=HTTPVerb(sVERB,sURL)
ReadHTTPVerb=SendAsyncRequest(oTransport,reqDATA)
Set ReadHTTPVerb=Nothing
End Function
Function SendHTTPVerb(ByVal sVERB,ByVal sURL,ByVal reqDATA)
Dim oTransport
Set oTransport=HTTPVerb(sVERB,sURL)
SendHTTPVerb=SendAsyncRequestWithData(oTransport,reqDATA)
Set oTransport=Nothing
End Function
Function ReqHTTPVerb(ByVal sVERB,ByVal sURL,ByVal reqDATA)
ReqHTTPVerb=SendHTTPVerb(sVERB,sURL,reqDATA)
End Function
Const FWX_HTTPDel_Verb="DELETE"
Const FWX_HTTPDel_Async=True
Function HTTPDel(ByVal sURL):Set HTTPDel=OpenHTTPVerb(FWX_HTTPDel_Verb,sURL,FWX_HTTPDel_Async,Null,Null):End Function
Function ASyncHTTPDel(ByVal sURL):Set ASyncHTTPDel=OpenHTTPDel(sURL,True,Null,Null):End Function
Function SyncHTTPDel(ByVal sURL):Set SyncHTTPDel=OpenHTTPDel(sURL,False,Null,Null):End Function
Function ReadHTTPDel(ByVal sURL):ReadHTTPDel=ReqHTTPVerb(FWX_HTTPDel_Verb,sURL,""):End Function
Function SendHTTPDel(ByVal sURL,ByVal reqDATA):SendHTTPDel=ReqHTTPVerb(FWX_HTTPDel_Verb,sURL,reqDATA):End Function
Function OpenHTTPDel(ByVal sURL,ByVal ASyncMode,ByVal sUSR,ByVal sPWD):Set OpenHTTPDel=OpenHTTPVerb(FWX_HTTPDel_Verb,sURL,ASyncMode,sUSR,sPWD):End Function
Const FWX_HTTPPut_Verb="PUT"
Const FWX_HTTPPut_Async=True
Function HTTPPut(ByVal sURL):Set HTTPPut=OpenHTTPVerb(FWX_HTTPPut_Verb,sURL,FWX_HTTPPut_Async,Null,Null):End Function
Function ASyncHTTPPut(ByVal sURL):Set ASyncHTTPPut=OpenHTTPPut(sURL,True,Null,Null):End Function
Function SyncHTTPPut(ByVal sURL):Set SyncHTTPPut=OpenHTTPPut(sURL,False,Null,Null):End Function
Function ReadHTTPPut(ByVal sURL):ReadHTTPPut=ReqHTTPVerb(FWX_HTTPPut_Verb,sURL,""):End Function
Function SendHTTPPut(ByVal sURL,ByVal reqDATA):SendHTTPPut=ReqHTTPVerb(FWX_HTTPPut_Verb,sURL,reqDATA):End Function
Function OpenHTTPPut(ByVal sURL,ByVal ASyncMode,ByVal sUSR,ByVal sPWD):Set OpenHTTPPut=OpenHTTPVerb(FWX_HTTPPut_Verb,sURL,ASyncMode,sUSR,sPWD):End Function
Const FWX_HTTPPost_Verb="POST"
Const FWX_HTTPPost_Async=True
Function HTTPPost(ByVal sURL):Set HTTPPost=OpenHTTPVerb(FWX_HTTPPost_Verb,sURL,FWX_HTTPPost_Async,Null,Null):End Function
Function ASyncHTTPPost(ByVal sURL):Set ASyncHTTPPost=OpenHTTPPost(sURL,True,Null,Null):End Function
Function SyncHTTPPost(ByVal sURL):Set SyncHTTPPost=OpenHTTPPost(sURL,False,Null,Null):End Function
Function ReadHTTPPost(ByVal sURL):ReadHTTPPost=ReqHTTPVerb(FWX_HTTPPost_Verb,sURL,""):End Function
Function SendHTTPPost(ByVal sURL,ByVal reqDATA):SendHTTPPost=ReqHTTPVerb(FWX_HTTPPost_Verb,sURL,reqDATA):End Function
Function OpenHTTPPost(ByVal sURL,ByVal ASyncMode,ByVal sUSR,ByVal sPWD):Set OpenHTTPPost=OpenHTTPVerb(FWX_HTTPPost_Verb,sURL,ASyncMode,sUSR,sPWD):End Function
Function PostData(ByRef sURL,ByVal reqDATA)
Dim objXMLHTTP
Set objXMLHTTP=HTTPPost(sURL)
PostData=SendAsyncRequestWithData(objXMLHTTP,reqDATA)
Set objXMLHTTP=Nothing
End Function
Function GetHTTPPost(ByRef sURL,ByVal reqDATA)
GetHTTPPost=PostData(objXMLHTTP,reqDATA)
End Function
Const FWX_HTTPGet_Verb="GET"
Const FWX_HTTPGet_Async=True
Function HTTPGet(ByVal sURL):Set HTTPGet=OpenHTTPVerb(FWX_HTTPGet_Verb,sURL,FWX_HTTPGet_Async,Null,Null):End Function
Function ASyncHTTPGet(ByVal sURL):Set ASyncHTTPGet=OpenHTTPGet(sURL,True,Null,Null):End Function
Function SyncHTTPGet(ByVal sURL):Set SyncHTTPGet=OpenHTTPGet(sURL,False,Null,Null):End Function
Function ReadHTTPGet(ByVal sURL):ReadHTTPGet=ReqHTTPVerb(FWX_HTTPGet_Verb,sURL,""):End Function
Function SendHTTPGet(ByVal sURL,ByVal reqDATA):SendHTTPGet=ReqHTTPVerb(FWX_HTTPGet_Verb,sURL,reqDATA):End Function
Function OpenHTTPGet(ByVal sURL,ByVal ASyncMode,ByVal sUSR,ByVal sPWD):Set OpenHTTPGet=OpenHTTPVerb(FWX_HTTPGet_Verb,sURL,ASyncMode,sUSR,sPWD):End Function
Const FWX_HTTPRequest_Verb="GET"
Const FWX_HTTPRequest_Async=True
Function HTTPRequest(ByVal sURL):Set HTTPRequest=OpenHTTPGet(sURL,FWX_HTTPRequest_Async,Null,Null):End Function
Function ASyncHTTPRequest(ByVal sURL):Set ASyncHTTPRequest=OpenHTTPGet(sURL,True,Null,Null):End Function
Function SyncHTTPRequest(ByVal sURL):Set SyncHTTPRequest=OpenHTTPGet(sURL,False,Null,Null):End Function
Function OpenHTTPRequest(ByVal sURL,ByVal ASyncMode,ByVal sUSR,ByVal sPWD):Set OpenHTTPRequest=OpenHTTPGet(sURL,ASyncMode,sUSR,sPWD):End Function
Function ExecuteRequest(ByRef objXMLHTTP)
ExecuteRequest=SendHTTPRequest(objXMLHTTP,FWX_HTTPRequest_Async)
End Function
Function SendRequest(ByRef objXMLHTTP)
SendRequest=ExecuteRequest(objXMLHTTP)
End Function
Function SendHTTPRequest(ByRef objXMLHTTP,ByVal async)
j("_STAT-HTTPRequest")=VOZ(j("_STAT-HTTPRequest"))+1
If async Then
SendHTTPRequest=SendAsyncRequest(objXMLHTTP)
Else
SendHTTPRequest=SendSyncRequest(objXMLHTTP)
End If
End Function
Function SendSyncRequest(ByRef objXMLHTTP)
Err.Clear():On Error Resume Next
objXMLHTTP.Send()
Dim o
o=objXMLHTTP.ResponseText
If IsTest Then j("___last_response")=o
SendSyncRequest=o
Set objXMLHTTP=Nothing
End Function
Function SendAsyncRequest(ByRef objXMLHTTP)
SendAsyncRequest=SendAsyncRequestWithData(objXMLHTTP,"")
End Function
Function PostAsyncRequest(ByRef objXMLHTTP,ByVal reqDATA)
PostAsyncRequest=SendAsyncRequestWithData(objXMLHTTP,reqDATA)
End Function
Function SendAsyncRequestWithData(ByRef objXMLHTTP,ByVal reqDATA)
Err.Clear():On Error Resume Next
Call CheckClientIsConnected()
If IsObject(reqDATA)Then
reqDATA=DIC2Qry(reqDATA)
End If
If reqDATA<>"" Then
objXMLHTTP.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
objXMLHTTP.setRequestHeader "Content-Length",Len(reqDATA)
End If
objXMLHTTP.Send reqDATA
If objXMLHTTP.readyState<>4 then
Dim xml_timeout
xml_timeout=VOV(j("FWX.HTTP.Timeout"),10)
objXMLHTTP.waitForResponse xml_timeout
End If
If objXMLHTTP.readyState<>4 then
o="REQUEST FAILED"&Enclose(LastHTTPRequest,", url=","")&Enclose(objXMLHTTP.Status,", status=","")&Enclose(objXMLHTTP.readyState,", readyState=","")&Enclose(objXMLHTTP.ResponseText,", response=","")&""
Else
If Err.Number<>0 then
o="LOCAL REQUEST ERROR"&Enclose(Err.Number,", error-number=","")&Enclose(Err.Description,", error-description=","")&Enclose(LastHTTPRequest,", url=","")&Enclose(objXMLHTTP.Status,", status=","")&Enclose(objXMLHTTP.readyState,", readyState=","")&Enclose(objXMLHTTP.ResponseText,", response=","")&""
Else
If objXMLHTTP.readyState=4 And objXMLHTTP.Status=200 Then
o=objXMLHTTP.responseBody
If LenB(o)>0 Then
o=BinaryToString(o)
Else
o=objXMLHTTP.ResponseText
End If
If IsTest Then j("___last_response")=o
Else
Select Case objXMLHTTP.Status
Case 404:
Case 301:
Case Else
o="REMOTE REQUEST ERROR"&Enclose(LastHTTPRequest,", url=","")&Enclose(objXMLHTTP.Status,", status=","")&Enclose(objXMLHTTP.readyState,", readyState=","")&Enclose(objXMLHTTP.ResponseText,", response=","")&""
End Select
objXMLHTTP.Abort
End If
End If
End If
Set objXMLHTTP=Nothing
SendAsyncRequestWithData=o
End Function
Function GetHumanResponse(ByVal sURL)
If sURL<>"" Then
Else
Exit Function
End If
Call CheckClientIsConnected()
Dim objXMLHTTP:Set objXMLHTTP=HTTPRequest(sURL)
objXMLHTTP.SetRequestHeader "User-Agent","	Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.3) Gecko/2008092417 Firefox/3.0.3"
objXMLHTTP.SetRequestHeader "Accept","text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
objXMLHTTP.SetRequestHeader "Accept-Language","en-gb,en;q=0.5"
objXMLHTTP.SetRequestHeader "Accept-Charset","UTF-8,ISO-8859-1;q=0.7,*;q=0.7"
objXMLHTTP.SetRequestHeader "Keep-Alive","300"
objXMLHTTP.SetRequestHeader "Connection","keep-alive"
If Request.Cookies<>"" Then
objXMLHTTP.SetRequestHeader "Cookies",Request.Cookies&""
End If
Dim o
o=ExecuteRequest(objXMLHTTP)
GetHumanResponse=o
End Function
Function GetResponseLazy(ByVal sURL)
GetResponseLazy=GetResponseToClient(sURL)
End Function
Function GetResponseAjax(ByVal sURL):GetResponseAjax=GetResponseToClient(sURL):End Function
Function GetResponseViaAjax(ByVal sURL):GetResponseViaAjax=GetResponseToClient(sURL):End Function
Function GetResponseToClient(ByVal sURL)
GetResponseToClient="<a class='auto-replace' href='"&sURL&"'>Loading...</a>"
End Function
Function GetResponseViaHTTPRequest(ByVal sURL):GetResponseViaHTTPRequest=GetResponseToServer(sURL):End Function
Function GetResponseToServer(ByVal sURL)
If sURL="" Then Exit Function
Call CheckClientIsConnected()
Dim objXMLHTTP
Set objXMLHTTP=HTTPRequest(sURL)
objXMLHTTP.SetRequestHeader "Content-Type","text/plain"
objXMLHTTP.SetRequestHeader "Accept","text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
objXMLHTTP.SetRequestHeader "Accept-Charset","utf-8,ISO-8859-1;q=0.7,*;q=0.7"
objXMLHTTP.SetRequestHeader "X-Requested-With","HTTPRequest"
GetResponseToServer=ExecuteRequest(objXMLHTTP)
End Function
Function GetResponse(ByVal sURL)
GetResponse=GetResponseToServer(sURL)
End Function
Function GetResponseAuthenticated(ByVal sURL,ByVal sUSR,ByVal sPWD)
If sURL="" Then Exit Function
Dim objXMLHTTP
Set objXMLHTTP=OpenHTTPRequest(sURL,FWX_HTTPRequest_Async,sUSR,sPWD)
GetResponseAuthenticated=ExecuteRequest(objXMLHTTP)
End Function
Function GetRawResponse(ByVal sURL)
GetRawResponse=ExecuteRequest(HTTPRequest(sURL))
End Function
Function GetResponseBody(ByVal sURL)
Call CheckClientIsConnected()
Dim objXMLHTTP:Set objXMLHTTP=SyncHTTPRequest(sURL)
objXMLHTTP.SetRequestHeader "Content-Type","text/plain"
objXMLHTTP.Send()
GetResponseBody=objXMLHTTP.ResponseBody
Set objXMLHTTP=Nothing
End Function
Function GetResponseBodyAuthenticated(ByVal sURL,ByVal sUSR,ByVal sPWD)
Call CheckClientIsConnected()
Dim objXMLHTTP
Set objXMLHTTP=OpenHTTPRequest(sURL,False,sUSR,sPWD)
objXMLHTTP.Send()
GetResponseBodyAuthenticated=objXMLHTTP.ResponseBody
Set objXMLHTTP=Nothing
End Function
Function Str2Xml(ByVal s):Set Str2Xml=LoadXml(s):End Function
Function LoadXml(ByVal s)
Call CheckClientIsConnected()
If IsObject(s)Then
Else
End If
If IsObject(s)Then
Set LoadXml=s
Else
Set LoadXml=XMLDOM()
LoadXml.Async=False
If s<>"" Then
Else
Exit Function
End If
If rxTest(s,"^https?://")Then
s=GetResponse(s)
Else
End If
s=rxReplace(s,"<\?xml\-stylesheet[^>]+>","")
s=rxReplace(s,"<feed xmlns=","<feed xmlns:atom=")
s=rxReplace(s,"[\s\n\r]+<","<")
s=Replace(s,"<?xml version=""1.0""?>","<?xml version=""1.0"" encoding=""utf-16""?>")
s=Replace(s,"encoding=""UTF-8""","encoding=""UTF-16""")
If Left(s,1)="<" Then
s=HTMLEntities_Remove(s)
LoadXml.LoadXml(s)
End If
End If
End Function
Function FlatXml(ByVal s)
Call CheckClientIsConnected()
Dim o
Set o=LoadXml(s)
If o.xml<>"" Then
o=rxReplace(o.xml&"","(</?)(\w+)\:(\w+)","$1$2_$3")
Set o=LoadXml(o)
End If
Set FlatXml=o
End Function
Function XmlList(ByVal x,ByVal n,ByVal t)
XmlList=XmlListWithLimit(x,n,t,0)
End Function
Function XmlListWithLimit(ByVal x,ByVal n,ByVal t,ByVal max)
Call CheckClientIsConnected()
Dim o:Set o=New QuickString
Dim oXML
Set oXML=FlatXml(x)
Dim node,nodes,i:i=1
Set nodes=oXML.selectNodes(n)
For Each node In nodes
If max<=0 Or i<=max Then
o.Add XmlMerge(node,t)
End If
i=i+1
Next
nodes="":
XmlListWithLimit=o
End Function
Function XmlMerge(ByVal x,ByVal t)
Call CheckClientIsConnected()
Dim o:o=""
Dim oXML:Set oXML=LoadXml(x)
o=MergeTemplate(t,oXML)
Set oXML=Nothing
XmlMerge=o
End Function
Function IsXml(ByRef x)
On Error Resume Next
IsXml=False
If IsObject(x)Then
Err.Clear()
IsXml=CBool(x.xml<>"")
IsXml=CBool(Err.Number=0)
If Err.Number<>0 Then Err.Clear()
End If
End Function
Function XmlFormat(x,strXSLFile)
Call CheckClientIsConnected()
Dim oXML:Set oXML=LoadXml(x)
Dim oXSL:Set oXSL=LoadXml(strXSLFile)
XmlFormat=oXML.transformNode(oXSL)
Set oXML=Nothing
Set oXSL=Nothing
End Function
Function ValidXPath(xml,xpath)
Err.Clear():On Error Resume Next
Set nodes=xml.selectNodes(xpath)
If nodes.Length>0 Then
ValidXPath=xpath
End If
End Function
Function RS2XMLDOM(ByRef rD)
Dim oXML:Set oXML=XMLDOM()
If Exists(rD)Then
rD.Save oXML,1
End If
Set RS2XMLDOM=oXML
End Function
Function RS2XML(ByRef rD)
Dim oXML
Set oXML=RS2XMLDOM(rD)
RS2XML=oXML.Xml
End Function
Function RS2NiceXML(ByRef rD)
Dim o:Set o=New QuickString
o.Add "<data>"
o.Add "<date>"&Now()&"</date>"
If Exists(rD)Then
Do While Exists(rD)
o.Add "<record>"
For Each f In rD.Fields
o.Add "<"&f.Name&"><![CDATA["&f.Value&"]]></"&f.Name&">"
Next
o.Add "</record>"
rD.MoveNext
Loop
rD.MoveFirst
End If
o.Add "</data>"
RS2NiceXML=o
End Function
Function WalkXml(ByVal root)
If root.childNodes.Length>0 Then
Dim o:o=""
o=o&"<ul>"
For Each node In root.childNodes
If Left(node.nodeName,1)="#" Then
Else
o=o&"<li>"&node.nodeName&""
o=o&WalkXml(node)&""
o=o&"</li>"
End If
Next
o=o&"</ul>"
Else
End If
Do While InStr(1,o,"<ul></ul>")>0
o=Replace(o,"<ul></ul>","")
Loop
WalkXml=o
End Function
Function Feed2JS(root,node):Feed2JS=Feed2JSON(root,node):End Function
Function Feed2JSON(root,node)
Set nodes=root.selectNodes(node)
Dim dat:Set dat=New QuickString
dat.Add "["'/*Feed2JSON-"& nodes.Length & "*/"
For Each node In nodes
dat.Add XML2JS(node)&","
Next
dat.Add "]"
dat=Replace(dat,",}","}")
dat=Replace(dat,",]","]")
Feed2JSON=dat
End Function
Function XML2JS(ByVal root):XML2JS=XML2JSON(root):End Function
Function XML2JSON(ByVal root)
If Not IsObject(root)Then Set root=LoadXml(root)
Dim dat
If root.childNodes.Length>0 Then
If root.childNodes.Length=1 And root.childNodes(0).NodeName="#text" Then
dat=""""&EncodeJS(root.childNodes(0).Text)&""""
Else
Set dat=New QuickString
dat.Add "{"
For Each node In root.childNodes
If Left(node.nodeName,1)="#" Then
Else
dat.Add """"&node.nodeName&""":"
dat.Add XML2JS(node)&","
End If
Next
dat.Add "}"
End If
Else
dat=""""&EncodeJS(root.Text)&""""
End If
dat=Replace(dat,",}","}")
XML2JSON=dat
End Function
Const xDC="@C@"
Const xDR="$R$"
Function RS2TXT_(ByRef rD,ByVal n)
Dim dat:Set dat=New QuickString
Dim f
If Exists(rD)Then
For Each f In rD.Fields
dat=dat&f.Name&xDC
Next
dat=Left(dat,Len(dat)-Len(xDC))
dat=dat&xDR
dat=dat&rD.GetString(,n,xDC,xDR,"")
dat=Replace(dat,xDC&xDR,xDR)
dat=Left(dat,Len(dat)-Len(xDR))
End If
RS2TXT_=dat
End Function
Function RS2TXTn(ByRef rD,ByVal n)
Dim dat:dat=RS2TXT_(rD,n)
dat=EncodeTSV(dat)
dat=Replace(dat,xDC,vbTab)
dat=Replace(dat,xDR,vbCRLf)
RS2TXTn=dat
End Function
Function RS2TXT(ByRef rD)
RS2TXT=RS2TXTn(rD,-1)
End Function
Function RS2CSVn(ByRef rD,ByVal n)
Dim dat
dat=RS2TXTn(rD,n)
dat=Replace(dat,"""","""""")
dat=""""&dat&""""
dat=rxReplace(dat,vbCRLf,""""&vbCRLf&"""")
dat=Replace(dat,vbTab,""",""")
RS2CSVn=dat
End Function
Function RS2CSV(ByRef rD)
RS2CSV=RS2CSVn(rD,-1)
End Function
Function RS2Table(ByRef rD)
Dim dat,th1,th2
dat=RS2TXT(rD)
dat=rxReplace(dat,"^([^\n\r]+)[\r\n]+([\s\S]+)","<thead><tr><th>$1</th></tr></thead><tbody><tr><td>$2</td></tr></tbody>")
dat=Replace(dat,vbTab,"</td><td>")
dat=Replace(dat,vbCRLf,"</td></tr>"&vbCRLf&"<tr><td>")
th1=rxSelect(dat,"^.*</thead>")
th2=rxReplace(th1,"<(/)?td>","<$1th>")
dat=rxReplace(dat,"^.*</thead>",th2)
dat=Replace(dat,"<td></td>","<td>-</td>")
dat=rxReplace(dat,"><(\w)",">"&vbCRLf&"<$1")
dat="<table>"&dat&"</table>"
RS2Table=dat
End Function
Function RS2TSV(ByRef rD)
Dim dat:dat=""""&RS2TXT(rD)&""""
dat=Replace(dat,vbCRLf,""""&vbCRLf&"""")
dat=Replace(dat,vbTab,""""&vbTab&"""")
RS2TSV=dat
End Function
Function JSON2CSV(ByVal s):JSON2CSV=JSON2TSV(s):End Function
Function JSON2TSV(ByVal s)
Dim arr
Select Case VarType(s)
Case 8
arr=STR2Array(s)
Case 8192
arr=s
Case Else
Exit Function
End Select
Dim csv
Set csv=New QuickString:csv.Divider=vbCRLf
Dim hed,row,col,fld
Set hed=New QuickString:hed.Divider=vbTab
For Each row In arr
Dim line
Set line=New QuickString:line.Divider=vbTab
For Each col In row
If csv.Length=0 Then
hed.Add col
End If
line.Add EncodeTSV(row(col))
Next
If csv.Length=0 Then
csv.Add hed
End If
csv.Add line
Next
JSON2TSV=csv
End Function
Function JSON2DIC(ByVal node)
Dim d
Set d=Dictionary()
TransferJSONIntoDIC node,d
Set JSON2DIC=d
End Function
Sub TransferJSONIntoDIC(ByVal node,ByRef d)
TransferJSONIntoDICWithPrefix node,d,""
End Sub
Sub TransferJSONIntoDICWithPrefix(ByVal node,ByRef d,ByVal prfx)
For Each key In node.keys()
d(prfx&key)=node.get(key)
If IsObject(node.get(key))Then
TransferJSONIntoDICWithPrefix node.get(key),d,prfx&key&"."
End If
Next
End Sub
Function RS2JSON(ByRef rD):RS2JSON=RS2JS(rD):End Function
Function RS2JS(ByRef rD)
Dim f
Dim dat:Set dat=New QuickString
dat.Add "["'/*RS2JS-"
If Exists(rD)Then
Do While Exists(rD)
dat.Add ""
dat.Add "{"
For Each f In rD.Fields
dat.Add """"&f.Name&""":"
Select Case RSFieldType(f)
Case "boolean":
dat.Add IIf(IsTrue(f.Value),"true","false")&""
Case "number":
dat.Add VOZ(f.value)&""
Case "date":
dat.Add """"&f.value&""""
Case Else:
dat.Add """"&EncodeJSON(Decode(f.Value))&""""
End Select
dat.Add ","
Next
dat.Add "},"
rD.MoveNext
Loop
rD.MoveFirst
End If
dat.Add "]"
dat=Replace(dat,",}","}")
dat=Replace(dat,",]","]")
RS2JS=dat
End Function
Function ADOXSL(ByRef rD,ByVal sXSL)
Dim o:o=""
Dim oXSL:Set oXSL=LoadXml(sXSL)
If oXSL.XML<>"" Then
Dim oXML:Set oXML=XMLDOM()
rD.Save oXML,1
o=oXML.transformNode(oXSL)
End If
ADOXSL=o
Set oXML=Nothing
Set oXSL=Nothing
End Function
Function XmlNode(ByRef xml,ByVal fld)
Dim obj
If fld="" Then
Set obj=xml
Else
If IsObject(obj)Then
If obj Is Nothing Then obj=""
End If
Err.Clear()
If Not IsObject(obj)Then Set obj=xml.selectSingleNode(".//"&fld)
If obj Is Nothing Then obj=""
Err.Clear()
If Not IsObject(obj)Then Set obj=xml.selectSingleNode(fld)
If obj Is Nothing Then obj=""
Err.Clear()
If Not IsObject(obj)Then Set obj=xml.selectSingleNode("//"&fld)
If obj Is Nothing Then obj=""
Err.Clear()
End If
Err.Clear()
If Not IsObject(obj)Then Set obj=Nothing
Set XmlNode=obj
End Function
Function XmlNodeText(ByRef xml,ByVal fld)
Dim o
Dim obj
Set obj=XmlNode(xml,fld)
If obj Is Nothing Then
Else
o=obj.Text
End If
obj=""
XmlNodeText=o
End Function
Function XmlNodeXml(ByRef xml,ByVal fld)
Dim o
Dim obj
Set obj=XmlNode(xml,fld)
If obj Is Nothing Then
Else
o=obj.xml
End If
obj=""
XmlNodeXml=o
End Function
Function XmlNodeAttr(ByRef xml,ByVal fld,ByVal atr)
Dim nod
Dim obj
Set nod=XmlNode(xml,fld)
If nod Is Nothing Then
Else
Set obj=nod.Attributes.getNamedItem(atr)
End If
nod=""
If Not IsObject(obj)Then
Set obj=Nothing
End If
Set XmlNodeAttr=obj
End Function
Function XmlNodeAttrText(ByRef xml,ByVal fld,ByVal atr)
Dim o
Dim obj
Set obj=XmlNodeAttr(xml,fld,atr)
If obj Is Nothing Then
Else
o=obj.Value&""
End If
obj=""
XmlNodeAttrText=o
End Function
Function XmlNodeValue(ByVal xml,ByVal tag)
If Not IsObject(xml)Then
Set xml=LoadXML(xml)
End If
If Not IsObject(xml)Then
Exit Function
End If
Dim nodes,s
If xml.NodeName=tag Then
s=xml.Text
Else
Set nodes=xml.getElementsByTagName(tag)
If nodes.Length>0 Then
s=nodes(0).Text
Else
s=XmlNodeText(xml,tag)
End If
End If
XmlNodeValue=s
End Function
Function NodeValue(ByVal xml,ByVal tag)
NodeValue=XmlNodeValue(xml,tag)
End Function
Dim FWX_MAIL__STEP_i
FWX_MAIL__STEP_i=0
Function FWX_MAIL__STEP()
FWX_MAIL__STEP_i=FWX_MAIL__STEP_i+1
FWX_MAIL__STEP=FWX_MAIL__STEP_i
End Function
Const cdoSendUsingPickup=1
Const cdoSendUsingPort=2
Const cdoAnonymous=0
Const cdoBasic=1
Const cdoNTLM=2
Const cdoContentDisposition="urn:schemas:mailheader:content-disposition"
Dim ma_r
Function mailer()
If Not IsObject(ma_r)Then Set ma_r=New fwxMailer
Set mailer=ma_r
End Function
Class fwxMailer
Public Property Get FromName():FromName=Me.Data("from.name"):End Property
Public Property Let FromName(ByVal s):Me.Data("from.name")=s:End Property
Public Property Get From():From=Me.Data("from"):End Property
Public Property Let From(ByVal s):Me.Data("from")=s:End Property
Public Property Get Recipients():Recipients=Me.Data("email"):End Property
Public Property Let Recipients(ByVal s):Me.Data("email")=s:End Property
Public Property Get Subject():Subject=Me.Data("subject"):End Property
Public Property Let Subject(ByVal s):Me.Data("subject")=s:End Property
Public Property Get Format():Format=Me.Data("format"):End Property
Public Property Let Format(ByVal s):Me.Data("format")=s:End Property
Public Property Get Body():Body=Me.Data("body"):End Property
Public Property Let Body(ByVal s):Me.Data("body")=s:End Property
Public Property Get BCC():BCC=Me.Data("bcc"):End Property
Public Property Let BCC(ByVal s):Me.Data("bcc")=s:End Property
Public Property Get CC():CC=Me.Data("cc"):End Property
Public Property Let CC(ByVal s):Me.Data("cc")=s:End Property
Public Property Get ReplyTo():ReplyTo=Me.Data("replyto"):End Property
Public Property Let ReplyTo(ByVal s):Me.Data("replyto")=s:End Property
Public Property Get Attachments():Attachments=Me.Data("attachments"):End Property
Public Property Let Attachments(ByVal s):Me.Data("attachments")=s:End Property
Public Data
Public Property Let Config(ByVal s)
UpdateCollection Me.Data,s
End Property
Public Sub Configure(ByVal c)
UpdateCollection Me.Data,c
End Sub
Public Sub ConfigureWithPrefix(ByVal c,ByVal p)
Dim f,x
For Each f In c
x=Mid(f,Len(p)+1)
If Left(f,Len(p))=p Then
If IsObject(c(f))Then
Set Me.Data(x)=c(f)
Else
Me.Data(x)=c(f)
End If
End If
Next
End Sub
Private Sub Class_Initialize()
Me.Reset()
End Sub
Public Sub Reset()
Set Data=Server.CreateObject("Scripting.Dictionary")
End Sub
Function SendCol(ByVal col):SendCol=False
Me.Configure col
SendCol=Me.Send()
End Function
Public Sending
Function Send():Send=False
If Me.Sending=True Then
Exit Function
Else
Me.Sending=True
End If
If IsTrue(FWX__Emails_Disabled)Then
FWX_AddHeader "FWX__Emails_Disabled","**** DISABLED ****"
DBG.Warning "FWX__Emails_Disabled"
DBG "MAIL: ################################### EMAILS ARE DISABLED"
Send=False
Exit Function
End If
Me.Data("key")=Me.Data("_key")
If Me.Data("from")="" Then Me.Data("from")=Me.Data("sender")
If Me.Data("email")="" Then Me.Data("email")=Me.ParseRecipients(Me.Data("to"))
If Me.Data("email")="" Then Me.Data("email")=Me.ParseRecipients(Me.Data("dest"))
If Me.Data("email")="" Then Me.Data("email")=Me.ParseRecipients(Me.Data("recipient"))
If Me.Data("email")="" Then Me.Data("email")=Me.ParseRecipients(Me.Data("recipients"))
If Me.Data("replyto")="" Then Me.Data("replyto")=Me.ParseRecipients(Me.Data("Reply-To"))
If Me.Data("replyto")="" Then Me.Data("replyto")=Me.ParseRecipients(Me.Data("reply-to"))
If Me.Data("subject")="" Then Me.Data("subject")=Me.Data("subj")
If Me.Data("subject")="" Then Me.Data("subject")=Me.Data("topic")
If Me.Data("body")="" Or GTZ(Me.Data("body"))Then Me.Data("body")=Me.Data("message")
If Me.Data("body")="" Or GTZ(ME.data("body"))Then Me.Data("body")=Me.Data("template")
If Me.Data("body")="" Or GTZ(ME.data("body"))Then Me.Data("body")=Me.Data("html")
If Me.Data("body")="" Or GTZ(ME.data("body"))Then Me.Data("body")=Me.Data("text")
Me.Data("utm_source")=TOT(Me.Data("utm_source"),"fyneworks")
Me.Data("utm_medium")=TOT(Me.Data("utm_medium"),"email")
Me.Data("utm_campaign")=TOT(Me.Data("utm_campaign"),TOT(Me.Data("_key"),Me.Data("_id")))
Me.Data("utm_content")=TOT(Me.Data("utm_content"),"")
Me.Data("utm_term")=TOT(Me.Data("utm_term"),"")' this is like Me.Data("utm_keyword")
Me.Data("body")=GA_TagLinks(Me.Data("body"),Me.Data)
Me.Data("body")=FixHTML_UserTemplate(Me.Data("body"))
Dim bot_email_tag
bot_email_tag=rxSelect(UAID,"Testomato|Uptimebot|GoogleBot")
If IsRobot Then bot_email_tag="bot"&Enclose(bot_email_tag,".","")
If Me.Data("max-width")="" Then
Me.Data("max-width")="600px"
End If
Me.Data("body")="<div id='fwx-mail' style='max-width:"&Me.Data("max-width")&";width:"&Me.Data("max-width")&"; magin:0 auto;'>"&vbCRLf&Enclose(parser.Parse(TemplateFile("mail/@header")),"<div id='fwx-mail-head'>","</div>")&vbCRLf&Enclose(Me.Data("body"),"<div id='fwx-mail-body'>","</div>")&vbCRLf&Enclose(parser.Parse(TemplateFile("mail/@footer")),"<div id='fwx-mail-foot'>","</div>")&vbCRLf&"</div>"&vbCRLf&"<div style='font-size:7.5pt;color:#ccc;'>"&vbCRLf&"fwx.mail."&FmtDate(Now(),"%Y%M%D.%H%N")&vbCRLf&Enclose(MyDomainID," / ","")&vbCRLf&Enclose(Me.Data("_key")," / ","")&vbCRLf&Enclose(Me.Data("_id")," / ","")&vbCRLf&Enclose(Me.Data("utm_campaign")," / ","")&vbCRLf&Enclose(bot_email_tag," / ","")&vbCRLf&"</div>"&vbCRLf&""
Me.Data("body")=rxReplace(Me.Data("body"),"(href|src|action)\=(""|')//([^/]+)/","$1=$2"&TOT(PROTOCOL,"http")&"://$3/")
Me.Data("body")=rxReplace(Me.Data("body"),"(href|src)\=(""|')\/","$1=$2http://"&MyWWWDomain&"/")
Me.Data("body")=Replace(Me.Data("body"),"http://"&SERVER_NAME&"/","http://"&MyWWWDomain&"/")
Me.Data("body")=Replace(Me.Data("body"),"http://localhost/","http://"&MyWWWDomain&"/")
Me.Data("body")=Replace(Me.Data("body"),"-fixed."&Mydomain,"."&Mydomain)
Me.Data("body")=rxReplace(Me.Data("body"),"(https?)://fixed-(sub(domain)-)?","$1://")
Me.Data("email")=Me.ParseRecipients(Me.Data("email"))
If Me.Data("email")="" Then Exit Function
If IsObject(App)Then
If Me.Data("from")="" Then Me.Data("from")=App.CFG("site.email")
If Me.Data("from")="" Then Me.Data("from")=App.CFG("shop.email")
End If
If IsObject(FWX)Then
If Me.Data("from")="" Then Me.Data("from")=FWX.Setting("email.from")
If Me.Data("from")="" Then Me.Data("from")=FWX.Setting("site.email")
If Me.Data("from")="" Then Me.Data("from")=FWX.Setting("shop.email")
End If
If Me.Data("from")="" Then Me.Data("from")="no-reply@"&Replace(DOMAIN_NAME,"www.","")&""
If InStr(1,Me.Data("from"),"<")>0 Then
Else
If IsObject(App)Then
If Me.Data("from.name")="" Then Me.Data("from.name")=App.CFG("name")
If Me.Data("from.name")="" Then Me.Data("from.name")=App.CFG("site.name")
If Me.Data("from.name")="" Then Me.Data("from.name")=App.CFG("shop.name")
End If
If IsObject(FWX)Then
If Me.Data("from.name")="" Then Me.Data("from.name")=FWX.Setting("name")
If Me.Data("from.name")="" Then Me.Data("from.name")=FWX.Setting("site.name")
If Me.Data("from.name")="" Then Me.Data("from.name")=FWX.Setting("shop.name")
End If
If Me.Data("from.name")<>"" Then
Me.Data("from")=Me.Data("from.name")&" <"&Me.Data("from")&">"
End If
End If
If IsObject(FWX)Then
Me.Data("bcc")=Enclose(Me.Data("bcc"),"",";")&FWX.Setting("email.bcc")
End If
If IsObject(FWX)Then
Me.Data("cc")=Enclose(Me.Data("cc"),"",";")&FWX.Setting("email.cc")
End If
Me.Data("format")="html"
Me.Data("cc")=Me.ParseRecipients(Me.Data("cc"))
Me.Data("bcc")=Me.ParseRecipients(Me.Data("bcc"))
Dim s
For Each s In Split("SMTP,Address,Port,Username,Password,SSL",",")
If s<>"SMTP" Then
s="SMTP."&s
End If
If Me.Data(s)="" Then
If IsObject(FWX)Then
DBG "MAIL: MAIL >>> LOAD SMTP SETTING '"&s&"' from FWX..."
Me.Data(s)=FWX.Setting(s)
End If
End If
If Me.Data(s)="" Then
DBG "MAIL: MAIL >>> LOAD SMTP SETTING '"&s&"' from Application.Contents..."
Me.Data(s)=Application.Contents(s)
End If
DBG "MAIL: MAIL >>> LOADED SMTP SETTING >>> "&s&"='"&Me.Data(s)&"'"
Next
If IsObject(mail)Then
If InStr(1,Me.Data("email"),MyDomain)>0 Or rxTest(Me.Data("email"),"fyneworks")Then
Else
If(IsBack Or IsStaff)And Me.Data("account")=MyAccount Then Me.Data("account")=0
For Each f In Split("account,invoice,quote,template",",")
If Not VOZ(Me.Data(f))>0 Then Me.Data(f)=VOZ(Query.Data(f))
If(IsBack Or IsStaff)And Me.Data("account")=MyAccount Then Me.Data("account")=0
If Not VOZ(Me.Data(f))>0 Then Me.Data(f)=VOZ(Query.Fata(f))
If(IsBack Or IsStaff)And Me.Data("account")=MyAccount Then Me.Data("account")=0
If Not VOZ(Me.Data(f))>0 Then Me.Data(f)=VOZ(j(f))
If(IsBack Or IsStaff)And Me.Data("account")=MyAccount Then Me.Data("account")=0
Next
If(IsBack Or IsStaff)And j("account")=MyAccount Then Me.Data("account")=0
If Me.Data("template")>0 Then
Me.Data("name")=TOT(Me.Data("shotName"),Me.Data("name"))
If Me.Data("name")="" Then Me.Data("name")=Me.Data("subject")
If Me.Data("context")="" Then Me.Data("context")=Me.Data("_id")
If Me.Data("request")="" Then Me.Data("request")=FULLREQUEST
If InStr(1,Me.Data("request"),"bulk")>0 Then Me.Data("request")="bulkmailer"
With mail.Shots
.ID=0
.ID=.Save(Me.Data)
Me.Data("shot")=.ID
End With
j("shot")=Me.Data("shot")
End If
If GTZ(Me.Data("template"))Or GTZ(Me.Data("account"))Or GTZ(Me.Data("quote"))Or GTZ(Me.Data("invoice"))Then
Me.Data("action")=App.Actions.Log("Email: "&TOT(Me.Data("subject"),Me.Data("name")),Me.Data)
j("action")=Me.Data("action")
End If
Me.Data("analytics")=parser.Parse(TemplateFile("mail/analytics"))
Me.Data("body")=Me.Data("body")&Me.Data("analytics")
End If
End If
DBG "MAIL: >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
DBG "MAIL: SEND EMAIL:"&vbCRLf
DBG "MAIL: from: '"&Encode(Me.Data("from"))&"'"&vbCRLf
DBG "MAIL: email: '"&Encode(Me.Data("email"))&"'"&vbCRLf
DBG "MAIL: subject: '"&Encode(Me.Data("subject"))&"'"&vbCRLf
DBG "MAIL: cc: '"&Encode(Me.Data("cc"))&"'"&vbCRLf
DBG "MAIL: bcc: '"&Encode(Me.Data("bcc"))&"'"&vbCRLf
DBG "MAIL: replyto: '"&Encode(Me.Data("replyto"))&"'"&vbCRLf
DBG "MAIL: format: '"&Encode(Me.Data("format"))&"'"&vbCRLf
DBG "MAIL: analytics: '"&Encode(Me.Data("analytics"))&"'"&vbCRLf
DBG "MAIL: body: '"&Encode(Me.Data("body"))&"'"&vbCRLf
DBG "MAIL: SMTP.SSL: '"&Encode(Me.Data("SMTP.SSL"))&"'"&vbCRLf '="y"
DBG "MAIL: SMTP.Port: '"&Encode(Me.Data("SMTP.Port"))&"'"&vbCRLf '="465"
DBG "MAIL: SMTP.Address: '"&Encode(Me.Data("SMTP.Address"))&"'"&vbCRLf '="smtp.gmail.com"
DBG "MAIL: SMTP.Username: '"&Encode(Me.Data("SMTP.Username"))&"'"&vbCRLf '="postman@fyneworks.com"
DBG "MAIL: SMTP.Password: '"&Encode(Me.Data("SMTP.Password"))&"'"&vbCRLf '="0gjdg*&%^oxs1GjI98CxZ[[_)l2cs2jh"
DBG "MAIL: />>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
DBG "MAIL: />>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
DBG "MAIL: />>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
Dim b,o,method_tried
b=False
For Each o In Split("CDOSYS,AspMail,CDONTS",",")
If IsTrue(b)Then
Else
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"",o&" start..."
DBG "MAIL: MAIL >>> TRY TO SEND EMAIL WITH "&o
Select Case o
Case "CDOSYS":
DBG "MAIL: MAIL >>> INVOKING CDOSYS."
b=Me.CDOSYS()
DBG "MAIL: MAIL >>> FINISHED TRYING CDOSYS."
Case "AspMail":
DBG "MAIL: MAIL >>> INVOKING AspMail."
b=Me.AspMail()
DBG "MAIL: MAIL >>> FINISHED TRYING AspMail."
Case "CDONTS":
DBG "MAIL: MAIL >>> INVOKING CDONTS."
b=Me.CDONTS()
DBG "MAIL: MAIL >>> FINISHED TRYING CDONTS."
End Select
DBG "MAIL: MAIL >>> TRY TO SEND EMAIL WITH "&o&" = '"&b&"'"
If IsTrue(b)Then
DBG "MAIL: MAIL >>> EMAIL SENT WITH "&o
DBG "MAIL: MAIL >>> EMAIL SENT WITH "&o
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"",o&" sent |F:"&Me.Data("from")&"|T:"&Me.Data("email")&"|CC:"&Me.Data("cc")&"|BCC:"&Me.Data("bcc")&"|S:"&Me.Data("subject")&"|B: "&Len(Me.Data("body"))&"b|FMT:"&Me.Data("format")&"|"&Now()
Else
DBG "MAIL: MAIL >>> EMAIL NOT SENT WITH "&o
DBG "MAIL: MAIL >>> EMAIL NOT SENT WITH "&o
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"",o&" not sent"
End If
End If
Next
DBG "MAIL: >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
DBG "MAIL: RESET!!!!"&vbCRLf
DBG "MAIL: >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
Me.Reset()
DBG "MAIL: >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
DBG "MAIL: SEND EMAIL DONE!"&vbCRLf
DBG "MAIL: >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
Me.Sending=False
Send=b
End Function
Function ParseRecipients(ByVal s)
s=rxReplace(s,"[\s\t\r\n;,]+",";")
s=rxReplace(s,";$","")
Dim x,y:y=""
For Each x In Split(s,";")
x=rxReplace(x,"[\s\t\r\n;,]+","")
If x<>"" And IsEmail(x)Then
If Not InStr(1,";"&y&";",";"&x&";")>0 Then
If y<>"" Then y=y&";"
y=y&x
End If
End If
Next
ParseRecipients=y
End Function
Function INIT_CDONTS()
Err.Clear():On Error Resume Next
Set INIT_CDONTS=Nothing
Set INIT_CDONTS=Server.CreateObject("CDONTS.NewMail")
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","INIT_CDONTS (CDONTS.NewMail) = "&Err.Number&" .:. "&IIf(Err.Number<>0,"error","OK!")
End Function
Public Function CDONTS():CDONTS=False
End Function
Function INIT_CDOSYS()
Err.Clear():On Error Resume Next
Set INIT_CDOSYS=Nothing
Set INIT_CDOSYS=Server.CreateObject("CDO.Message")
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","INIT_CDOSYS (CDO.Message) = "&Err.Number&" .:. "&IIf(Err.Number<>0,"error","OK!")
End Function
Function CDOSYS():CDOSYS=False
FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","CDOSYS START"
FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","CDOSYS INIT CDO.Message"
Dim oCDOSYSCon,oCDOSYSMail
Set oCDOSYSMail=Me.INIT_CDOSYS()
If IsObject(oCDOSYSMail)Then
If oCDOSYSMail Is Nothing Then
DBG "CDOSYS INIT FAILED"
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","oCDOSYSMail  = NOTHING, QUIT! "&Err.Description
Err.Clear()
Exit Function
Else
DBG "CDOSYS INIT WORKED"
End If
Else
DBG "CDOSYS INIT NOT EVEN AN OBJECT"
End If
Dim s
For Each s In Split("SMTP,Address,Port,Username,Password,SSL",",")
If s<>"SMTP" Then
s="SMTP."&s
End If
If Me.Data(s)="" Then
If IsObject(FWX)Then
DBG "MAIL: MAIL >>> CDOSYS >>> LOAD SMTP SETTING '"&s&"' from FWX..."
Me.Data(s)=FWX.Setting(s)
End If
End If
If Me.Data(s)="" Then
DBG "MAIL: MAIL >>> CDOSYS >>> LOAD SMTP SETTING '"&s&"' from Application.Contents("&SERVER_NAME&"-"&s&")..."
Me.Data(s)=Application.Contents(SERVER_NAME&"-"&s)
End If
If Me.Data(s)="" Then
DBG "MAIL: MAIL >>> CDOSYS >>> LOAD SMTP SETTING '"&s&"' from Application.Contents..."
Me.Data(s)=Application.Contents(s)
End If
DBG "MAIL: MAIL >>> CDOSYS >>> LOADED SMTP SETTING >>> "&s&"='"&Me.Data(s)&"'"
Next
If Me.Data("SMTP")="" Then
If Not StartsWith(SERVER_IP,"109.")Then
FWX_AddHeader " FWX-MAIL-DISABLED","Missing SMTP Settings (CDOSYS)"
DIE "Missing SMTP Settings (CDOSYS)"
DBG "MAIL: MAIL >>> CDOSYS >>> QUITTING. MISSING SMTP SETTINGS"
Exit Function
End If
End If
Set oCDOSYSCon=Server.CreateObject("CDO.Configuration")
With oCDOSYSCon
If Me.Data("SMTP")<>"" Then
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")=TOT(Me.Data("SMTP.Address"),"smtp.gmail.com")
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=TOT(Me.Data("SMTP.Port"),465)
.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")=Me.Data("SMTP.Username")
.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")=Me.Data("SMTP.Password")
.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")=IIf(Me.Data("SMTP.SSL"),True,False)
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")=90
Else
.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")="127.0.0.1"'localhost"
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")=90
End If
.Fields.Update()
End With
Set oCDOSYSMail.Configuration=oCDOSYSCon
With oCDOSYSMail
.Subject=Me.Data("subject")
.From=Me.Data("from")
.CC=Me.Data("cc")
.BCC=Me.Data("bcc")
If Me.Data("replyto")<>"" Then
.ReplyTo=Me.Data("replyto")
Else
.ReplyTo=Me.Data("from")
End If
Dim sAT
If Me.Data("attachments")<>"" Then
If IsNumeric(Me.Data("attachments"))Then
Me.Data("attachments")=""
ElseIf GTZ(Me.Data("attachments"))Then
Me.Data("attachments")=""
End If
End If
If Me.Data("attachments")<>"" Then
Dim attachments,attachment_name,attachment_path,attachment_url,attachment_data,attachment_file
For Each sAT In Split(Me.Data("attachments"),";")
sAT=TOE(sAT)
If IsNumeric(sAT)Or GTZ(sAT)Then
sAT=""
End If
If sAT<>"" Then
attachments=VOZ(attachments)+1
attachment_name=rxSelect(sAT,"[^\\\/~]+$")
attachment_path=rxReplace(sAT,"~[^~]+$","")
DBG "MAIL: EMAIL - ATTACHMENT: '"&attachment_name&"' > "&attachment_path
If FileExists(attachment_path)Then
DBG "MAIL: EMAIL - ATTACHMENT - PHYSICAL FILE: '"&PhysicalPath(attachment_path)&"'"
.AddAttachment PhysicalPath(attachment_path)
Else
attachment_url=FullyQualifiedURL(attachment_path)
.AddAttachment attachment_url
End If
With .Attachments(attachments).Fields
.Item(cdoContentDisposition)="attachment;filename="&attachment_name
.Update
End With
End If
Next
End If
If 1=0 Then
ElseIf Me.Data("format")="html" Then
.HTMLBody=Me.Data("body")'"<html><body>html message</body></html>"
ElseIf Me.Data("format")="url" Then
.CreateMHTMLBody Me.Data("body")'"http://www.paulsadowski.com/wsh/"
ElseIf Me.Data("format")="file" Then
.CreateMHTMLBody Me.Data("body")'"file://c|/temp/test.htm"
Else
.TextBody=Me.Data("body")'"plain text message"
End If
Dim sTO
For Each sTO In Split(Me.Data("email"),";")
.To=sTO
.Fields.Update()
DBG "MAIL: MAIL >>> CDOSYS >>> SEND!"
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","CDOSYS SEND!"
DBG ""
If Me.CanSend(sTO)Then .Send()
DBG "MAIL: MAIL >>> CDOSYS >>> SENT!"
DBG ""
Next
End With
Set oCDOSYSMail=Nothing
Set oCDOSYSCon=Nothing
CDOSYS=True
End Function
Function INIT_AspMail()
Err.Clear():On Error Resume Next
Set INIT_AspMail=Nothing
Set INIT_AspMail=Server.CreateObject("Persits.MailSender")
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","INIT_AspMail (Persits.MailSender) = "&Err.Number&" .:. "&IIf(Err.Number<>0,"error","OK!")
End Function
Function AspMail():AspMail=False
Dim oAspMail
Set oAspMail=Me.INIT_AspMail()
If oAspMail Is Nothing Then
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","oAspMail  = NOTHING, QUIT! "&Err.Description
Err.Clear()
Exit Function
End If
Dim s
For Each s In Split("SMTP,Address,Port,Username,Password,SSL",",")
If s<>"SMTP" Then
s="SMTP."&s
End If
If Me.Data(s)="" Then
If IsObject(FWX)Then
DBG "MAIL: MAIL >>> AspMail >>> LOAD SMTP SETTING '"&s&"' from FWX..."
Me.Data(s)=FWX.Setting(s)
End If
End If
If Me.Data(s)="" Then
DBG "MAIL: MAIL >>> AspMail >>> LOAD SMTP SETTING '"&s&"' from Application.Contents("&SERVER_NAME&"-"&s&")..."
Me.Data(s)=Application.Contents(SERVER_NAME&"-"&s)
End If
If Me.Data(s)="" Then
DBG "MAIL: MAIL >>> AspMail >>> LOAD SMTP SETTING '"&s&"' from Application.Contents..."
Me.Data(s)=Application.Contents(s)
End If
DBG "MAIL: "
DBG "MAIL: MAIL >>> AspMail >>> LOADED SMTP SETTING >>> "&s&"='"&Me.Data(s)&"'"
DBG "MAIL: "
Next
If Me.Data("SMTP")="" Then
If Not StartsWith(SERVER_IP,"109.")Then
FWX_AddHeader " FWX-MAIL-DISABLED","Missing SMTP Settings (ASPMAIL)"
DIE "Missing SMTP Settings (ASPMAIL)"
DBG "MAIL: MAIL >>> AspMail >>> QUITTING. MISSING SMTP SETTINGS"
Exit Function
End If
End If
If Me.Data("SMTP")<>"" Then
With oAspMail
.Host=Me.Data("SMTP.Address")
.Username=Me.Data("SMTP.Username")
.Password=Me.Data("SMTP.Password")
.From=Me.Data("from")
.FromName=Me.Data("from")
If Me.Data("replyto")<>"" Then
.AddReplyTo Me.Data("replyto")
Else
.AddReplyTo Me.Data("from")
End If
.Subject=Me.Data("subject")
.Body=Me.Data("body")
.IsHTML=True
Dim sTO
For Each sTO In Split(Me.Data("bcc"),";")
.AddBCC sTO
Next
For Each sTO In Split(Me.Data("cc"),";")
.AddCC sTO
Next
For Each sTO In Split(Me.Data("email"),";")
If Me.CanSend(sTO)Then .AddAddress sTO
Next
DBG "MAIL: MAIL >>> ASPMAIL >>> SEND!"
If(IsTest Or IsAdmin)Then FWX_AddHeader " FWX-MAIL-"&FWX_MAIL__STEP()&"","ASPMAIL SEND!"
DBG ""
.Send()
DBG ""
DBG "MAIL: MAIL >>> ASPMAIL >>> SENT."
End With
End If
Set oAspMail=Nothing
oAspMail=""
AspMail=CBool(Err.Number=0)
End Function
Function CanSend(ByVal s)
If IsLocal Then
Else
If IsLive Then
CanSend=True
Else
If EndsWith(s,MyDomain)Then
CanSend=True
Else
If IsTest Then
DBG.Warning "Not sending to '"&s&"' (test mode)"
Else
DBG.FatalError "Can't send email to '"&s&"'."
End If
End If
End If
End If
End Function
End Class
Function fwxSendEmail(sFrom,sRecipients,sSubject,sBody,sFormat)
With mailer
.From=sFrom
.Recipients=sRecipients
.Subject=sSubject
.Format=sFormat
.Body=sBody
.Send()
End With
fwxSendEmail=CBool(Not(Err.number<>0))
End Function
Function SendTextEmail(sFrom,sTO,sSubject,sBody)
SendTextEmail=fwxSendEmail(sFrom,sTO,sSubject,sBody,"text")
End Function
Function SendRichEmail(sFrom,sTO,sSubject,sBody)
SendRichEmail=fwxSendEmail(sFrom,sTO,sSubject,sBody,"html")
End Function
Function SendEmail(sFrom,sTO,sSubject,sBody)
SendEmail=SendRichEmail(sFrom,sTO,sSubject,sBody)
End Function
Function H404_m()
If IsStaffPC Or IsStaff Or Application.Contents("DisableEmailTracking")<>"" Or InStr(1,REFERRER,Mydomain)>0 Or InStr(1,REFERRER,SERVER_NAME)>0 Or InStr(1,Application.Contents("IgnoreStatsFrom"),VISITOR_IP)>0 Then
DIE "no"
End If
RHTM
j("referrer")=Request.ServerVariables("HTTP_REFERER")
j("shot")=rxSelect(Query.FULLID,"\d+$")
If Not GTZ(j("shot"))Then DIE ""
j("name")=rxReplace(Query.FULLID,"^m\.|\.\d+$","")
k=j("name")&":"&j("shot")
If MySession(k)<>"" Then
DIE "ok"
End If
MySession(k)="y"
Select Case j("name")
Case "o":j("name")="open"
Case "r":j("name")="reply"
Case "p":j("name")="print"
Case "f":j("name")="forward"
End Select
j("keyword")=Query.FQ("keyword,q,key,text")
If Not IsObject(mail)Then Set mail=New fwxMAIL
j("ping")=mail.Shots.pings.Save(j)
If Not IsObject(App)Then Set App=New fwxAPP
UpdateCollection j,mail.Shots.FindFields("[subject],[account],[invoice],[_key]","shot="&j("shot"))
j("key")=j("_key")
j("is_service")=1
j("bucket")="service"
j("e")="mail.Shots"
j("notes")=j("subject")
j("action-gentle")="y"
j("action")=App.Actions.Log(PCase(j("name"))&" Email",j)
DIE j("ping")&":"&j("action")
End Function
Dim MyEncryptionKey
MyEncryptionKey=Trim(Right(RemoveSymbols(Server.MapPath("/")),5)&Left(Reverse(RemoveSymbols(Request.ServerVariables("HTTP_HOST"))),5))
Function Encrypt(ByVal s)
Encrypt=DoEncrypt(MyEncryptionKey,s)
End Function
Function Decrypt(ByVal s)
Decrypt=DoDecrypt(MyEncryptionKey,s)
End Function
Function StrongEncrypt(ByVal s)
StrongEncrypt=RC4(MyEncryptionKey,s)
End Function
Function StrongDecrypt(ByVal s)
StrongDecrypt=RC4(MyEncryptionKey,s)
End Function
Function OneWayEncrypt(ByVal s)
OneWayEncrypt=SHA256(s)
End Function
Function EncryptIt(ByVal sPWD,ByVal sMSG,ByVal iACT)
On Error Resume Next
sPWD=RemoveSymbols(sPWD)
Dim strAlphabet
Dim intMessageChar
Dim intCodewordChar
Dim intShiftAdjust
Dim intHomeLocation
strAlphabet="0123456789AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
intCodewordChar=1
For intMessageChar=1 To Len(sMSG)
If InStr(1,strAlphabet,Mid(sMSG,intMessageChar,1),vbBinaryCompare)>0 Then
If InStr(1,strAlphabet,Mid(sPWD,intCodewordChar,1),vbBinaryCompare)>0 Then
intShiftAdjust=InStr(1,strAlphabet,Mid(sPWD,intCodewordChar,1),vbBinaryCompare)
intHomeLocation=InStr(1,strAlphabet,Mid(sMSG,intMessageChar,1),vbBinaryCompare)
If iACT=0 Then intShiftAdjust=intHomeLocation-intShiftAdjust
If iACT=1 Then intShiftAdjust=intHomeLocation+intShiftAdjust
If intShiftAdjust>Len(strAlphabet)Then intShiftAdjust=intShiftAdjust-Len(strAlphabet)
If intShiftAdjust<1 Then intShiftAdjust=intShiftAdjust+Len(strAlphabet)
Else
intShiftAdjust=1
End If
EncryptIt=EncryptIt&Mid(strAlphabet,intShiftAdjust,1)
Else
EncryptIt=EncryptIt&Mid(sMSG,intMessageChar,1)
End If
If intCodewordChar>Len(sPWD)Then intCodewordChar=1 Else intCodewordChar=intCodewordChar+1
Next
End Function
Function DoEncrypt(ByVal sPWD,ByVal sMSG)
DoEncrypt=EncryptIt(sPWD,sMSG,0)
End Function
Function DoDecrypt(ByVal sPWD,ByVal sMSG)
DoDecrypt=EncryptIt(sPWD,sMSG,1)
End Function
Public Function Hash(ByVal text)
a=1
For i=1 to len(text)
a=sqr(a*i*asc(mid(text,i,1)))
Next
Rnd(-1)
Randomize a
For i=1 to 16
Hash=Hash&Chr(int(rnd*256))
Next
End function
Dim lcl_o_AN
Function ANNumber(ByVal s)
If Not IsObject(lcl_o_AN)Then Set lcl_o_AN=New AlphaNumericNumberParser
ANNumber=lcl_o_AN.ToNumber(s)
End Function
Function AlphaNumber(ByVal s)
AlphaNumber=ANNumber(s)
End Function
Function ANString(ByVal s)
If Not IsObject(lcl_o_AN)Then Set lcl_o_AN=New AlphaNumericNumberParser
ANString=lcl_o_AN.ToString(s)
End Function
Function AlphaID(ByVal s)
AlphaID=ANString(s)
End Function
Dim lcl_o_MD5
Function MD5(ByVal sMessage)
If TOE(sMessage)="" Then Exit Function
If Not IsObject(lcl_o_MD5)Then
Set lcl_o_MD5=New fwxMD5
End If
MD5=UCase(lcl_o_MD5.MD5(sMessage))
End Function
Dim lcl_o_SHA256
Function SHA256(ByVal sMessage)
If TOE(sMessage)="" Then Exit Function
If Not IsObject(lcl_o_SHA256)Then
Set lcl_o_SHA256=New fwxSHA256
End If
SHA256=lcl_o_SHA256.SHA256(sMessage)
End Function
Function HMAC_SHA256(ByVal secret,ByVal sMessage)
If TOE(sMessage)="" Then Exit Function
Dim wsc,sig
Set wsc=GetObject("script:"&Server.MapPath("/fwx/inc/sha256/sha256.wsc"))
wsc.hexcase=0
sig=wsc.b64_hmac_sha256(secret,sMessage)
Set wsc=Nothing
HMAC_SHA256=sig
End Function
Function SHA512(ByVal str)
Dim objMD5,objUTF8
Dim arrByte
Dim strHash
Set objUnicode=CreateObject("System.Text.UnicodeEncoding")
arrByte=objUnicode.GetBytes_4(str)
Set objUnicode=Nothing
Set objSHA512=Server.CreateObject("System.Security.Cryptography.SHA512Managed")
strHash=objSHA512.ComputeHash_2((arrByte))
Set objSHA512=Nothing
SHA512=ToBase64(strHash)
End Function
Function ToBase64(ByVal rabyt)
Dim xml:Set xml=CreateObject("MSXML2.DOMDocument.3.0")
xml.LoadXml "<root />"
xml.documentElement.dataType="bin.base64"
xml.documentElement.nodeTypedValue=rabyt
ToBase64=xml.documentElement.Text
Set xml=Nothing
End Function
Function HashedPassword(ByVal userPassword)
If StartsWith(userPassword,"md5-")Then
userPassword="key-"&SHA256(UCase(rxReplace(userPassword,"^md5\-","")))
ElseIf Not StartsWith(userPassword,"key-")Then
userPassword="key-"&SHA256(UCase(MD5(userPassword)))
End If
HashedPassword=userPassword
End Function
Function HashRecipe(ByVal userPassword,ByVal userSalt)
HashRecipe="["& HashedPassword(userPassword)&"+("&userSalt&")@"&MyDomain&"]^"&FWX_Security_Pepper
End Function
Function SeasonedHash(ByVal userPassword,ByVal userSalt)
Dim stretch:stretch=FWX_Security_Strech
Dim stint:stint=1
Dim base:base=HashRecipe(userPassword,userSalt)
Dim hash:hash=SHA256(base)
If stint<stretch Then
Do While stint<stretch
hash=SHA256(hash&" * "&userSalt&" * "&FWX_Security_Pepper)
stint=stint+1
Loop
End If
SeasonedHash=hash
End Function
Function SignMessage(ByVal message,ByVal salt)
SignMessage=SHA256(message&" w!t4 "&salt&" &aNd "&FWX_Security_Pepper)
End Function
Const FWX_Security_SignedMessageSalt="$igNed* mESs@ge# sa!t!/?!"
Class FWX_Security_SignedMessage__Parser
Public salt
Public expiry
Public token
Public stamp
Public signature
Public message
Private Sub Class_Initialize()
Me.stamp=FmtDate(Now(),fwxDateOrdinal)
If IsObject(FWX)Then
Me.salt=FWX.Setting("security.SignedMessageSalt")
Else
Me.Salt=FWX_Security_SignedMessageSalt
End If
End Sub
Public Property Let payload(ByVal s)
Me.token=rxReplace(s,"[^@]*\@(\d+)\.(.+)$","$1.$2")
If Me.token<>"" Then
Me.stamp=rxSelect(Me.token,"^\d+")
Me.signature=rxSelect(Me.token,"[^\.]+$")
Me.message=Replace(s,"@"&Me.token,"")
End If
End Property
Public Property Get Sign(ByVal s)
Sign=SignMessage(s,Me.salt)
End Property
Public Property Get body()
body=Me.message&"@"&Me.stamp
End Property
Public Property Get signed()
signed=Me.body&"."&Me.Sign(Me.body)
End Property
Public Default Property Get payload()
payload=Me.signed
End Property
Public Function isGenuine()
isGenuine=False
If Me.signature=Me.Sign(Me.body)Then
isGenuine=True
End If
End Function
Public Function isFresh()
isFresh=False
Dim d,a,e
a=Me.stamp
e=Me.expiry
If e="" Then
isFresh=True
Else
d=rxReplace(a,"^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\..+)?$","$1-$2-$3 $4:$5:$6")
If IsDate(d)Then
d=MakeDateFromString(d)
If Not IsDate(e)Then
If VOZ(e)<>0 Then
e=DateAdd("n",(0-Abs(VOZ(e))),Now())
Else
e=""
End If
End If
If e<>"" Then e=MakeDateFromString(e)
If IsDate(e)Then
If d>e Then
isFresh=True
End If
End If
End If
End If
End Function
Public Function isValid()
isValid=False
If Me.isFresh And Me.isGenuine Then
isValid=True
End If
End Function
Public Function accepted()
If Me.isValid Then accepted=Me.message
End Function
End Class
Function SignedMessageWithSalt(ByVal message,ByVal salt)
Dim m
Set m=New FWX_Security_SignedMessage__Parser
With m
If salt<>"" Then .salt=salt
.expiry=""
.message=message
SignedMessageWithSalt=.signed
End With
Set m=Nothing
End Function
Function SignedMessage(ByVal message)
SignedMessage=SignedMessageWithSalt(message,"")
End Function
Function TimedMessageWithSalt(ByVal payload,ByVal expiry,ByVal salt)
Dim m
Set m=New FWX_Security_SignedMessage__Parser
With m
If salt<>"" Then .salt=salt
.expiry=expiry
.payload=payload
TimedMessageWithSalt=.accepted
End With
Set m=Nothing
End Function
Function TimedMessage(ByVal payload,ByVal expiry)
TimedMessage=TimedMessageWithSalt(payload,expiry,"")
End Function
Function VerifiedMessageWithSalt(ByVal payload,ByVal salt)
VerifiedMessageWithSalt=TimedMessageWithSalt(payload,"",salt)
End Function
Function VerifiedMessage(ByVal payload)
VerifiedMessage=VerifiedMessageWithSalt(payload,"")
End Function
Function UUID()
Randomize Timer
Dim i,RndNum
For i=0 to 7
RndNum=CLng(rnd * "&HFFFF")
If i=3 Then RndNum=(RndNum And "&HFFF")Or "&H4000"
If i=4 Then RndNum=(RndNum And "&H3FFF")Or "&H8000"
UUID=UUID+String(4-Len(Hex(RndNum)),"0")+LCase(Hex(RndNum))
If i=1 Or i=2 Or i=3 Or i=4 Then UUID=UUID+"-"
Next
UUID="{"&UCase(UUID)&"}"
End Function
Function GUID()
GUID=Left(CreateObject("Scriptlet.TypeLib").Guid,38)
End Function
Const authTokenExp=86400
Const authTokenSalt="authTokenSalt"
Const FWX_Security_Strech_Default=1
Function FWX_Security_Strech()
FWX_Security_Strech=VOV(Application.Contents("security.stretch"),FWX_Security_Strech_Default)
End Function
Const FWX_Security_Pepper_Default="Sa!t thi$ way... W0o0! W@lk T4iS wa-hay!"
Function FWX_Security_Pepper()
FWX_Security_Pepper=TOT(Application.Contents("security.pepper"),FWX_Security_Salt_Default)
End Function
Function RC4(ByVal Key,ByVal Data)
dim KeyBytes(255)
dim CypherBytes(255)
KeyLen=Len(Key)
if KeyLen=0 Then Exit function
KeyText=Key
For i=0 To 255
KeyBytes(i)=Asc(Mid(Key,((i)Mod(KeyLen))+1,1))
Next
if Len(Data)=0 Then Exit function
For i=0 To 255
CypherBytes(i)=i
Next
Jump=0
For i=0 To 255
Jump=(Jump+CypherBytes(i)+KeyBytes(i))Mod 256
Tmp=CypherBytes(i)
CypherBytes(i)=CypherBytes(Jump)
CypherBytes(Jump)=Tmp
Next
i=0
Jump=0
For X=1 To Len(Data)
i=(i+1)Mod 256
Jump=(Jump+CypherBytes(i))Mod 256
T=(CypherBytes(i)+CypherBytes(Jump))Mod 256
Tmp=CypherBytes(i)
CypherBytes(i)=CypherBytes(Jump)
CypherBytes(Jump)=Tmp
RC4=RC4&Chr(Asc(Mid(Data,X,1))Xor CypherBytes(T))
Next
End function
Const b64Key="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Dim b64EncMap
Function b64Enc(ByVal i)
If Not IsArray(b64EncMap)Then
Redim b64EncMap(63)
Dim K:For K=0 To Len(b64Key)-1
b64EncMap(K)=Mid(b64Key,K+1,1)
Next
End If
b64Enc=b64EncMap(i)
End Function
Dim b64DecMap
Function b64Dec(ByVal i)
If Not IsArray(b64DecMap)Then
Redim b64DecMap(127)
Dim K:For K=0 To Len(b64Key)-1
b64DecMap(ASC(b64Enc(K)))=K
Next
End If
b64Dec=b64DecMap(i)
End Function
Function base64Encode(strPlain)
If Len(strPlain)=0 then
base64Encode=""
Exit Function
End If
iBy3=(Len(strPlain)\ 3)* 3
strReturn=""
iIndex=1
Do While iIndex<=iBy3
iFirst=Asc(Mid(strPlain,iIndex+0,1))
iSecond=Asc(Mid(strPlain,iIndex+1,1))
iiThird=Asc(Mid(strPlain,iIndex+2,1))
strReturn=strReturn&b64Enc((iFirst \ 4)AND 63)
strReturn=strReturn&b64Enc(((iFirst * 16)AND 48)+((iSecond \ 16)AND 15))
strReturn=strReturn&b64Enc(((iSecond * 4)AND 60)+((iiThird \ 64)AND 3))
strReturn=strReturn&b64Enc(iiThird AND 63)
iIndex=iIndex+3
Loop
If iBy3<Len(strPlain)then
iFirst=Asc(Mid(strPlain,iIndex+0,1))
strReturn=strReturn&b64Enc((iFirst \ 4)AND 63)
If(Len(strPlain)MOD 3)=2 then
iSecond=Asc(Mid(strPlain,iIndex+1,1))
strReturn=strReturn&b64Enc(((iFirst * 16)AND 48)+((iSecond \ 16)AND 15))
strReturn=strReturn&b64Enc((iSecond * 4)AND 60)
Else
strReturn=strReturn&b64Enc((iFirst * 16)AND 48)
strReturn=strReturn&"="
End If
strReturn=strReturn&"="
End If
base64Encode=strReturn
End Function
Function base64Decode(strEncoded)
If Len(strEncoded)=0 then
base64Decode=""
Exit Function
End If
iRealLength=Len(strEncoded)
Do While Mid(strEncoded,iRealLength,1)="="
iRealLength=iRealLength-1
Loop
Do While instr(strEncoded," ")<>0
strEncoded=left(strEncoded,instr(strEncoded," ")-1)&"+"&Mid(strEncoded,instr(strEncoded," ")+1)
Loop
strReturn=""
iBy4=(iRealLength\4)*4
iIndex=1
Do While iIndex<=iBy4
iFirst=b64Dec(Asc(Mid(strEncoded,iIndex+0,1)))
iSecond=b64Dec(Asc(Mid(strEncoded,iIndex+1,1)))
iThird=b64Dec(Asc(Mid(strEncoded,iIndex+2,1)))
iFourth=b64Dec(Asc(Mid(strEncoded,iIndex+3,1)))
strReturn=strReturn&chr(((iFirst * 4)AND 255)+((iSecond \ 16)AND 3))
strReturn=strReturn&chr(((iSecond * 16)AND 255)+((iThird \ 4)AND 15))
strReturn=strReturn&chr(((iThird * 64)AND 255)+(iFourth AND 63))
iIndex=iIndex+4
Loop
If iIndex<iRealLength then
iFirst=b64Dec(Asc(Mid(strEncoded,iIndex+0,1)))
iSecond=b64Dec(Asc(Mid(strEncoded,iIndex+1,1)))
strReturn=strReturn&chr(((iFirst * 4)AND 255)+((iSecond \ 16)AND 3))
If iRealLength MOD 4=3 then
iThird=b64Dec(Asc(Mid(strEncoded,iIndex+2,1)))
strReturn=strReturn&chr(((iSecond * 16)AND 255)+((iThird \ 4)AND 15))
End If
End If
base64Decode=strReturn
End Function
Function SimpleXor(strIn,strKey)
if len(strIn)=0 or len(strKey)=0 then
SimpleXor=""
exit function
end if
iInIndex=1
iKeyIndex=1
strReturn=""
do while iInIndex<=len(strIn)
strReturn=strReturn&chr(asc(mid(strIn,iInIndex,1))XOR asc(mid(strKey,iKeyIndex,1)))
iInIndex=iInIndex+1
if iKeyIndex=len(strKey)then iKeyIndex=0
iKeyIndex=iKeyIndex+1
loop
simpleXor=strReturn
End function
Function RandomKey(keyLength)
sDefaultChars="abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
iKeyLength=keyLength
iDefaultCharactersLength=Len(sDefaultChars)
Randomize
For iCounter=1 To iKeyLength
iPickedChar=Int((iDefaultCharactersLength * Rnd)+1)
sMyKey=sMyKey&Mid(sDefaultChars,iPickedChar,1)
Next
RandomKey=sMyKey
End Function
Sub rw(ByVal s)
Response.Write s
End Sub
Dim w_c:w_c=0
Function w(ByVal str)
w_c=w_c+1
Rw "<div style='"
If w_c MOD 2=0 Then rw "background:#eee;"
If w_c MOD 2>0 Then rw "background:#fff;"
If Err.Description<>"" Or Err.Number<>0 Then Rw "color:red;"
Rw "'"
If Err.Description<>"" Or Err.Number<>0 Then Rw " title="""&Err.Description&""""
Rw ">"&str&"</div>"&vbCRLf
End Function
Sub rPdf():rTYPE("application/pdf"):End Sub
Sub rJS():rTYPE("text/javascript"):End Sub
Sub rJSON():rTYPE("application/json"):End Sub
Sub rTxt():rTYPE("text/plain"):End Sub
Sub rCsv():rTYPE("text/csv"):End Sub
Sub rXml():rTYPE("text/xml"):End Sub
Sub rHtm():rTYPE("text/html"):End Sub
Sub rHtml():rTYPE("text/html"):End Sub
Dim FWX_ResponseType
FWX_ResponseType="html"
Sub rTYPE(ByVal s)
FWX_ResponseType=Right(s,Len(s)-InStr(1,s,"/"))
FWX_ContentType s
End Sub
Function DIE(ByVal s)
If IsObject(DBG)Then
DBG ""
DBG ""
End If
RW s
RE
End Function
Function DIETXT(ByVal s)
RTXT
DIE s
End Function
Function DIEJS(ByVal s)
RJSON
If IsObject(s)Then s=DIC2JSON(s)
DIE s
End Function
Function DIEJSON(ByVal s)
RJSON
If IsObject(s)Then s=DIC2JSON(s)
DIE s
End Function
Function DIEXML(ByVal s)
RXML
If IsObject(s)Then s=DIC2XML(s)
DIE s
End Function
Function DIEWrap(ByVal s,ByVal x,ByVal y)
DIE Enclose(s,x,y)
End Function
Function DIEasJSON(ByVal s)
DIEasJSON=DIEWrapJSON(s)
End Function
Function DIEWrapJSON(ByVal s)
If IsObject(s)Then s=DIC2JSON(s)
DIEJSON "{"&s&"}"
End Function
Function DIEasXML(ByVal s)
DIEasXML=DIEWrapXML(s)
End Function
Function DIEWrapXML(ByVal s)
If IsObject(s)Then s=DIC2XML(s)
DIEXML "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"&vbCRLf&"<xml>"&vbCRLf&s&vbCRLf&"</xml>"
End Function
Function ColRem(ByVal s,ByVal f)
On Error Resume Next
s.Remove(f)
If Err.Number>0 Then Err.Clear()
End Function
Function ColGet(ByVal s,ByVal f)
On Error Resume Next
ColGet=s(f)
If Err.Number>0 Then Err.Clear()
End Function
Const FWX__API_Format="json"
Function DIEWrapped(ByVal s,ByVal format)
If format<>"" Then
If Not rxTest(format,"(xml|js|json)")Then format=""
End If
If format="" Then
If format="" And IsObject(s)Then format=s("format")
If format="" And IsObject(j)Then format=j("format")
If format="" And IsObject(Query)Then format=Query.FQ("format")
If format="" And IsObject(Query)Then format=rxSelect(Query.EXT,"^(js|json|xml|txt|csv|tsv)$")
If format="" And IsObject(FWX)Then format=FWX.Setting("api.format")
If format="" Then format=FWX__API_Format
If IsObject(s)Then
If ColGet(s,"format")="" Then ColRem s,"format"
End If
End If
Select Case LCase(format)
Case "xml":
If IsObject(s)Then s=DIC2XML(s)
DIEWrapXML(s)
Case "json","js":
If IsObject(s)Then s=DIC2JSON(s)
DIEWrapJSON(s)
Case Else:
If IsObject(s)Then s=DIC2TXT(s)
DIETXT s
End Select
End Function
Function DBGvars(ByVal s)
DBG PrintColVars(j,s)
End Function
Function DIEDic(ByVal s,ByVal format)
DIEDic=DIEWrapped(s,format)
End Function
Function DIEColVars(ByVal c,ByVal v)
DIEDic FilterCol(c,v),""
End Function
Function DIEDicVars(ByVal c,ByVal v)
DIEDic FilterCol(c,v),""
End Function
Function DIEj(ByVal s,ByVal format)
DIEj=DIEWrapped(FilterJ(s),format)
End Function
Function DIEVars(ByVal v)
DIEj v,""
End Function
Function DIECol(ByVal c)
DIEDic c,""
End Function
Function DIC2TXT(ByVal c)
Dim n,o:Set o=New QuickString
For Each n In c
If IsObject(c(n))Then
o.Add n&"="&TOT(c(n),"[object]")&vbCRLf
Else
o.Add n&"="&c(n)&vbCRLf
End If
Next
DIC2TXT=o&""
End Function
Function DIEn(ByVal s)
j("status")="n"
j("error")=TOT(s,"")
DIEvars "status,error"
End Function
Function DIEerr(s):DIEerr=DIEn(s):End Function
Function DIEbad(s):DIEbad=DIEn(s):End Function
Function DIEerror(s):DIEerror=DIEn(s):End Function
Function DIEy(ByVal s)
j("status")="y"
j("message")=TOT(s,"")
DIEvars "status,message"
End Function
Function DIEok(s):DIEok=DIEn(s):End Function
Function DIEgood(s):DIEgood=DIEn(s):End Function
Function DIEsuccess(s):DIEsuccess=DIEn(s):End Function
Function DIEHTM(ByVal s)
RHTML
DIE s
End Function
Function DIEHTML(ByVal s)
RHTML
DIE s
End Function
Sub re():Response.End():End Sub
Sub rc():On Error Resume Next:Response.Clear():End Sub
Sub rst(ByVal s):On Error Resume Next:FWX_ResponseStatus s:End Sub
Function DebugCol(col)
Dim o
o="<table cellspacing=5>"
For Each Var In col
o=o&"<tr class=Small><td align=right class=Bold><b>"&Var&":</b></td><td valign=top>"&Encode(Slice(col(Var),50))&"</td></tr>"
Next
DebugCol=o&"</table>"
End Function
Function DescribeRS(ByRef rsDescribe)
For Each fld In rsDescribe.Fields
w fld.Name&"("&fld.Type&")"
Next
End Function
Function fwxStatus(ByVal s,ByVal m)
fwxSendEmail "status@"&MyDomain&"","dev@fyneworks.com",s,m,"html"
End Function
Function fwxError(ByVal s,ByVal m)
fwxSendEmail "error@"&MyDomain&"","dev@fyneworks.com",s,m,"html"
End Function
Function IsTrue(ByVal s)
IsTrue=False
s=Trim(LCase(s))
If s="" Then Exit Function
If InStr(1,s,"=")>0 Then
Dim ifa
ifa=Split(s,"=")
If UBound(ifa)=1 Then
If ifa(0)<>"" Then
s=IIf(ifa(0)=ifa(1),"true","")
End If
End If
If s="" Then Exit Function
End If
If s="true" Then IsTrue=True:Exit Function
If s="y" Then IsTrue=True:Exit Function
If s="1" Then IsTrue=True:Exit Function
If s="-1" Then IsTrue=True:Exit Function
If s="on" Then IsTrue=True:Exit Function
If s="ok" Then IsTrue=True:Exit Function
If s="yes" Then IsTrue=True:Exit Function
End Function
Function IsFalse(ByVal s)
IsFalse=CBool(Not IsTrue(s))
End Function
Function Num(ByVal n)
Err.Clear():On Error Resume Next
Num=CDbl(n)
If Err.Description<>"" Or Err.Number<>0 Then Num=0
Err.Clear()
End Function
Function VOV(ByVal n,ByVal d)
n=Num(n):d=Num(d)
If n=0 Then n=d
If n=0 Then n=0
VOV=n
End Function
Function VOZ(ByVal n):VOZ=VOV(n,0):End Function
Function GTZ(ByVal n):GTZ=CBool(VOZ(n)>0):End Function
Function LTZ(ByVal n):LTZ=CBool(VOZ(n)<0):End Function
Function GT1(ByVal n):GT1=CBool(VOZ(n)>1):End Function
Function TOT(ByVal a,ByVal b)
Err.Clear():On Error Resume Next
Dim s:s=""
If Not s<>"" Then s=Trim(a&"")
If Not s<>"" Then s=Trim(b&"")
TOT=s:Err.Clear()
End Function
Function TOE(ByVal s):TOE=TOT(s,""):End Function
Function OTOT(ByRef o,ByVal b)
On Error Resume Next
OTOT=o&""
If Err.Number<>0 Then
OTOT=b
Err.Clear()
End If
End Function
Function TextOrSpace(ByVal a)
a=TOE(a)
If a="" Then a=" "
TextOrSpace=a
End Function
Function IsEmpty(ByVal s)
Err.Clear():On Error Resume Next
IsEmpty=False
s=Trim(s)&""
IsEmpty=CBool(Not(Len(s)>0))
If Err.Description<>"" Or Err.Number<>0 Then IsEmpty=True
Err.Clear()
End Function
Function IsSet(ByVal s)
IsSet=Not IsEmpty(s)
End Function
Function IsEmail(ByVal s)
IsEmail=False
r.Pattern=rxEmail
If IsSet(s)Then
IsEmail=r.Test(s)
End If
End Function
Function IsPostcode(ByVal s)
IsPostcode=IsUKPostcode(s)
End Function
Function IsUKPostcode(ByVal s)
On Error Resume Next
IsUKPostcode=False
If IsSet(s)Then
r.Pattern=rxUKPostcode
IsUKPostcode=r.Test(s)
End If
End Function
Function IsPhone(ByVal s)
IsPhone=False
s=s&""
r.Pattern="[\-\+\.\,\(\)\ ]+"
s=r.Replace(s&"","")&""
If Not(Len(s)>=6)Then Exit Function
r.Pattern="^\d+$"
If Not(r.Test(s))Then Exit Function
IsPhone=True
End Function
Function IsUrl(ByVal s)
IsUrl=rxTest(s,rxURL)
End Function
Function IsAddress(rec)
If Not IsObject(rec)Then Exit Function
IsAddress=False
If IsEmpty(ColItem(rec,"address1"))Then Exit Function
If IsEmpty(ColItem(rec,"city"))Then Exit Function
If Not rxTest(ColItem(rec,"country"),"^(UA|AE|OM)$")Then
If IsEmpty(ColItem(rec,"postcode"))Then Exit Function
End If
If ColItem(rec,"country")="US" Then
If IsEmpty(ColItem(rec,"county"))Then Exit Function
End If
IsAddress=True
End Function
Function IsValidString(ByVal s)
Dim sInvalidChars
Dim bln
Dim i
sInvalidChars="!#$%^&*()=+{}[]|\;:/?>,<"
For i=1 To Len(sInvalidChars)
If InStr(s,Mid(sInvalidChars,i,1))>0 Then bln=True
If bln Then Exit For
Next
IsValidString=CBool(Not bln)
End Function
Function IsValidID(ByVal s)
bln=False
If IsEmpty(s)Then bTemp=True
If Not IsValidString(s)Then bln=True
If(InStr(s,".")>0)Then bln=True
If(InStr(s," ")>0)Then bln=True
If(Len(s)<>Len(Trim(s)))Then bln=True
IsValidID=Not bln
End Function
Function IsNothing(ByRef obj)
IsNothing=True
If Not IsObject(entFolders)Then
Exit Function
End If
If(entFolders Is Nothing)Then
Exit Function
End If
IsNothing=Not CBool(Err.Number=0)
End Function
Const adOpenForwardOnly=0
Const adOpenKeyset=1
Const adOpenDynamic=2
Const adOpenStatic=3
Const adHoldRecords=&H00000100
Const adMovePrevious=&H00000200
Const adAddNew=&H01000400
Const adDelete=&H01000800
Const adUpdate=&H01008000
Const adBookmark=&H00002000
Const adApproxPosition=&H00004000
Const adUpdateBatch=&H00010000
Const adResync=&H00020000
Const adNotify=&H00040000
Const adLockReadOnly=1
Const adLockPessimistic=2
Const adLockOptimistic=3
Const adLockBatchOptimistic=4
Const adRunAsync=&H00000010
Const adStateClosed=&H00000000
Const adStateOpen=&H00000001
Const adStateConnecting=&H00000002
Const adStateExecuting=&H00000004
Const adUseServer=2
Const adUseClient=3
Const adEmpty=0
Const adTinyInt=16
Const adSmallInt=2
Const adInteger=3
Const adBigInt=20
Const adUnsignedTinyInt=17
Const adUnsignedSmallInt=18
Const adUnsignedInt=19
Const adUnsignedBigInt=21
Const adSingle=4
Const adDouble=5
Const adCurrency=6
Const adDecimal=14
Const adNumeric=131
Const adBoolean=11
Const adError=10
Const adUserDefined=132
Const adVariant=12
Const adIDispatch=9
Const adIUnknown=13
Const adGUID=72
Const adDate=7
Const adDBDate=133
Const adDBTime=134
Const adDBTimeStamp=135
Const adBSTR=8
Const adChar=129
Const adVarChar=200
Const adLongVarChar=201
Const adWChar=130
Const adVarWChar=202
Const adLongVarWChar=203
Const adBinary=128
Const adVarBinary=204
Const adLongVarBinary=205
Const adFldMayDefer=&H00000002
Const adFldUpdatable=&H00000004
Const adFldUnknownUpdatable=&H00000008
Const adFldFixed=&H00000010
Const adFldIsNullable=&H00000020
Const adFldMayBeNull=&H00000040
Const adFldLong=&H00000080
Const adFldRowID=&H00000100
Const adFldRowVersion=&H00000200
Const adFldCacheDeferred=&H00001000
Const adEditNone=&H0000
Const adEditInProgress=&H0001
Const adEditAdd=&H0002
Const adEditDelete=&H0004
Const adRecOK=&H0000000
Const adRecNew=&H0000001
Const adRecModified=&H0000002
Const adRecDeleted=&H0000004
Const adRecUnmodified=&H0000008
Const adRecInvalid=&H0000010
Const adRecMultipleChanges=&H0000040
Const adRecPendingChanges=&H0000080
Const adRecCanceled=&H0000100
Const adRecCantRelease=&H0000400
Const adRecConcurrencyViolation=&H0000800
Const adRecIntegrityViolation=&H0001000
Const adRecMaxChangesExceeded=&H0002000
Const adRecObjectOpen=&H0004000
Const adRecOutOfMemory=&H0008000
Const adRecPermissionDenied=&H0010000
Const adRecSchemaViolation=&H0020000
Const adRecDBDeleted=&H0040000
Const adGetRowsRest=-1
Const adPosUnknown=-1
Const adPosBOF=-2
Const adPosEOF=-3
Const adBookmarkCurrent=0
Const adBookmarkFirst=1
Const adBookmarkLast=2
Const adMarshalAll=0
Const adMarshalModifiedOnly=1
Const adAffectCurrent=1
Const adAffectGroup=2
Const adAffectAll=3
Const adFilterNone=0
Const adFilterPendingRecords=1
Const adFilterAffectedRecords=2
Const adFilterFetchedRecords=3
Const adFilterPredicate=4
Const adSearchForward=1
Const adSearchBackward=-1
Const adPromptAlways=1
Const adPromptComplete=2
Const adPromptCompleteRequired=3
Const adPromptNever=4
Const adModeUnknown=0
Const adModeRead=1
Const adModeWrite=2
Const adModeReadWrite=3
Const adModeShareDenyRead=4
Const adModeShareDenyWrite=8
Const adModeShareExclusive=&Hc
Const adModeShareDenyNone=&H10
Const adXactUnspecified=&Hffffffff
Const adXactChaos=&H00000010
Const adXactReadUncommitted=&H00000100
Const adXactBrowse=&H00000100
Const adXactCursorStability=&H00001000
Const adXactReadCommitted=&H00001000
Const adXactRepeatableRead=&H00010000
Const adXactSerializable=&H00100000
Const adXactIsolated=&H00100000
Const adXactCommitRetaining=&H00020000
Const adXactAbortRetaining=&H00040000
Const adPropNotSupported=&H0000
Const adPropRequired=&H0001
Const adPropOptional=&H0002
Const adPropRead=&H0200
Const adPropWrite=&H0400
Const adErrInvalidArgument=&Hbb9
Const adErrNoCurrentRecord=&Hbcd
Const adErrIllegalOperation=&Hc93
Const adErrInTransaction=&Hcae
Const adErrFeatureNotAvailable=&Hcb3
Const adErrItemNotFound=&Hcc1
Const adErrObjectInCollection=&Hd27
Const adErrObjectNotSet=&Hd5c
Const adErrDataConversion=&Hd5d
Const adErrObjectClosed=&He78
Const adErrObjectOpen=&He79
Const adErrProviderNotFound=&He7a
Const adErrBoundToCommand=&He7b
Const adErrInvalidParamInfo=&He7c
Const adErrInvalidConnection=&He7d
Const adErrStillExecuting=&He7f
Const adErrStillConnecting=&He81
Const adParamSigned=&H0010
Const adParamNullable=&H0040
Const adParamLong=&H0080
Const adParamUnknown=&H0000
Const adParamInput=&H0001
Const adParamOutput=&H0002
Const adParamInputOutput=&H0003
Const adParamReturnValue=&H0004
Const adCmdUnknown=&H0008
Const adCmdText=&H0001
Const adCmdTable=&H0002
Const adCmdStoredProc=&H0004
Const adCmdFile=&H0100
Const adSchemaProviderSpecific=-1
Const adSchemaAsserts=0
Const adSchemaCatalogs=1
Const adSchemaCharacterSets=2
Const adSchemaCollations=3
Const adSchemaColumns=4
Const adSchemaCheckConstraints=5
Const adSchemaConstraintColumnUsage=6
Const adSchemaConstraintTableUsage=7
Const adSchemaKeyColumnUsage=8
Const adSchemaReferentialContraints=9
Const adSchemaTableConstraints=10
Const adSchemaColumnsDomainUsage=11
Const adSchemaIndexes=12
Const adSchemaColumnPrivileges=13
Const adSchemaTablePrivileges=14
Const adSchemaUsagePrivileges=15
Const adSchemaProcedures=16
Const adSchemaSchemata=17
Const adSchemaSQLLanguages=18
Const adSchemaStatistics=19
Const adSchemaTables=20
Const adSchemaTranslations=21
Const adSchemaProviderTypes=22
Const adSchemaViews=23
Const adSchemaViewColumnUsage=24
Const adSchemaViewTableUsage=25
Const adSchemaProcedureParameters=26
Const adSchemaForeignKeys=27
Const adSchemaPrimaryKeys=28
Const adSchemaProcedureColumns=29
Function CloneRS(ByRef source)
If Not IsObject(source)Then Exit Function
Set CloneRS=Recordset()
For Each fld In source.Fields
attr=fld.Attributes
If fld.Type>=200 And(Not attr>0)Then
attr=104
End If
CloneRS.Fields.Append fld.Name,fld.Type,attr
If Err.Description<>"" Or Err.Number<>0 Then
DBG "CloneRS *** error cloning field "&fld.Name&", "&fld.Type&", "&attr&""
Err.Clear
Else
DBG "CloneRS - clonned field "&fld.Name&", "&fld.Type&", "&attr&""
End If
Next
If Not CloneRS.State=1 Then CloneRS.Open
End Function
Function Connect(ByRef objRecordset,ByVal connStr)
Err.Clear():On Error Resume Next
If IsEmpty(connStr)Then Exit Function
Call CheckClientIsConnected()
If Not IsObject(objRecordset)Then
CreateRS objRecordset
End If
j("_STAT-Connect")=VOZ(j("_STAT-Connect"))+1
With objRecordset
.ActiveConnection=connStr
If Err.Description<>"" Or Err.Number<>0 Then
DBG.FatalError "DATA: ERROR #"&Err.Number&" ["&Err.Description&"]"&IIf(IsLocal,"<br/>Connection String: "&connStr&"","")&""
Err.Clear()
Connect=False
Else
Connect=True
End If
End With
End Function
Function ConnectRS(ByRef objRecordset,ByVal connStr)
ConnectRS=Connect(objRecordset,connStr)
End Function
Function CreateRS(ByRef objRecordset)
If IsObject(CreateRS)Then
DBG "FWX.Data.CreateRS> (already object)"
Else
DBG "FWX.Data.CreateRS> (already object)"
End If
Set objRecordset=Server.CreateObject("ADODB.Recordset")
With objRecordset
.CursorLocation=3
.CursorType=3
.LockType=3
CreateRS=True
End With
DBG "FWX.Data.CreateRS> (register with killer)"
RSKiller.Register objRecordset
DBG "FWX.Data.CreateRS> (done)"
End Function
Function MakeRS(ByVal fields)
Dim rs,s
Set rs=Server.CreateObject("ADODB.Recordset")
For Each s In Split(fields,",")
rs.Fields.Append s,adVarChar,500
Next
rs.Open
Set MakeRS=rs
End Function
Function CreateRS(ByRef objRecordset)
If IsObject(CreateRS)Then
DBG "FWX.Data.CreateRS> (already object)"
Else
DBG "FWX.Data.CreateRS> (already object)"
End If
Set objRecordset=Server.CreateObject("ADODB.Recordset")
With objRecordset
.CursorLocation=3
.CursorType=3
.LockType=3
CreateRS=True
End With
DBG "FWX.Data.CreateRS> (register with killer)"
RSKiller.Register objRecordset
DBG "FWX.Data.CreateRS> (done)"
End Function
Function MakeRS(ByVal fields)
Dim rs,s
Set rs=Server.CreateObject("ADODB.Recordset")
For Each s In Split(fields,",")
rs.Fields.Append s,adVarChar,500
Next
rs.Open
Set MakeRS=rs
End Function
Function FWXDB(ByVal s)
Dim a,x
If s<>"" Then
s=rxReplace(s,";\s+",";")
a=Split(s,";")
If UBound(a)=4 Then
x="DRIVER={SQL Server}; SERVER="&a(0)&"; DATABASE="&a(1)&"; UID="&a(2)&"; PWD="&a(3)&";"
End If
End If
FWXDB=x
End Function
Function DeleteRS(ByRef objRecordset)
blnTempConn=False
With objRecordset
blnTempConn=Not ValidConn(objRecordset)
If blnTempConn Then
ConnectRS objRecordset,FWX.ConnectionString
End If
.Delete()
If blnTempConn Then
Set .ActiveConnection=Nothing
End If
End With
End Function
Function RecordCount(ByRef rs)
RecordCount=RSCount(rs)
End Function
Function RSCount(ByRef rs)
On Error Resume Next
Dim d
d=rs.Recordcount
If Err.Number<>0 Then
Err.Clear()
RSCount=0
Else
RSCount=d
End If
End Function
Function IsRecordset(ByRef rs)
On Error Resume Next
Dim d
d=rs.Recordcount
If Err.Number<>0 Then
Err.Clear()
IsRecordset=False
Else
IsRecordset=True
End If
End Function
Function PrintRS(ByRef rs)
PrintRS=DescribeRS(rs)
End Function
Function DescribeRS(ByRef rsDescribe)
Err.Clear():On Error Resume Next
rsDescribe.MoveFirst()
If Err.Number<>0 Then
Err.Clear()
output=output&"<div>Not a recordset</div>"
DescribeRS=output
Exit Function
End If
output=output&"<br/>rs.State = "&rsDescribe.State
output=output&"<br/>rs.EOF = "&rsDescribe.EOF
output=output&"<br/>rs.BOF = "&rsDescribe.BOF
output=output&"<br/>rs.RecordCount = "&VOZ(rsDescribe.RecordCount)
output=output&"<br/>rs.Source = "&Replace(Encode(rsDescribe.Source),"[","[ ")
If Not Exists(rsDescribe)Then
output=output&"<div>empty recordset</div>"
DescribeRS=output
Exit Function
End If
output=output&"<table border=1  cellspacing=5>"
output=output&"<tr>"
For Each fld In rsDescribe.Fields
output=output&"<td>"&fld.Name&"("&fld.Type&")</td>"
Next
output=output&"</tr>"
rsDescribe.Filter=""
Do While Exists(rsDescribe)
output=output&"<tr>"
For Each fld In rsDescribe.Fields
output=output&"<td>"
output=output&Encode(Left(rsDescribe(fld.Name),100))
output=output&"</td>"
Next
output=output&"</tr>"
rsDescribe.MoveNext
Loop
rsDescribe.MoveFirst
output=output&"</table>"
DescribeRS=output
End Function
Function ExistsOrKill(ByRef lclRecordSet)
b=Exists(lclRecordSet)
ExistsOrKill=b
End Function
Function Exists(ByRef lclRecordSet)
Err.Clear():On Error Resume Next
Call CheckClientIsConnected()
Exists=False
If IsObject(lclRecordSet)Then
If lclRecordSet.State=1 Then
If(lclRecordSet.BOF Or lclRecordSet.EOF)Then
Else
If Err.Number=0 Then
Exists=True
End If
End If
End If
End If
End Function
Function ExistsXml(ByRef lclRecordSet)
ExistsXml=False
Call CheckClientIsConnected()
If Not IsObject(lclRecordSet)Then Exit Function
If lclRecordSet.State<>1 Then Exit Function
iRecordCount=lclRecordSet.RecordCount
If Err.Description<>"" Or Err.Number<>0 Then
iRecordCount=-1
End If
ExistsXml=CBool(iRecordCount>=0)
Err.Clear()
End Function
Set j("Queries")=Dictionary()
Function ReloadRS(ByRef rREC)
If IsObject(FWX)Then
If FWX.FSOOnly Then Exit Function
End If
Dim sRunSQL
sRunSQL=rREC.Source
KillRS rRec
Set rRec=RunSQL(sRunSQL)
Set ReloadRS=rRec
End Function
Function RunSQL(ByVal sRunSQL)
Dim rO
Call CreateRS(rO)
Call OpenRS(rO,sRunSQL)
Set RunSQL=rO
End Function
Function GetRS(ByVal sRunSQL)
Set GetRS=RunSQL(GetRS)
End Function
Function NewRS(ByVal sRunSQL)
Set NewRS=GetRS(sRunSQL)
End Function
Function ExeSQL(ByVal sRunSQL)
If IsObject(FWX)Then
If FWX.FSOOnly Then Exit Function
End If
KillRS RunSQL(sRunSQL)
End Function
Function KillRS(ByRef oRS)
On Error Resume Next
oRS.Close()
If Err.Number<>0 Then
Err.Clear()
Else
Set oRS=Nothing
End If
oRS=""
End Function
Dim LastSQL
Function OpenRS(ByRef objRecordset,ByVal sRunSQL)
Err.Clear():On Error Resume Next
DBG "FWX.Data.OpenRS> ..."
DBG "FWX.Data.OpenRS> FWX is object? "&isObject(FWX)
If IsObject(FWX)Then
DBG "FWX.Data.OpenRS> FWX.FSOOnly = "&FWX.FSOOnly
If FWX.FSOOnly Then Exit Function
End If
DBG "FWX.Data.OpenRS> CheckClientConnected..."
CheckClientConnected
DBG "FWX.Data.OpenRS> CheckClientConnected - proceed"
DBG "FWX.Data.OpenRS> objRecordset?"
If Not IsObject(objRecordset)Then
DBG "FWX.Data.OpenRS> objRecordset? No, create..."
CreateRS objRecordset
Else
DBG "FWX.Data.OpenRS> objRecordset? Yes, already an ibject"
End If
sRunSQL=TOT(sRunSQL,"")
sRunSQL=rxReplace(sRunSQL,"[\r\n\s]+"," ")
sRunSQL=rxReplace(sRunSQL,"\[([\r\n\s\/\\\?])","[")
If sRunSQL="" Then
If IsLocal Then
DBG.FatalError "Trying to execute empty SQL"
Else
DBG.Warning "Trying to execute empty SQL"
Exit Function
End If
Else
DBG "FWX.Data.OpenRS> sRunSQL="&sRunSQL
End If
LastSQL=sRunSQL
With objRecordset
DBG "FWX.Data.OpenRS> A) objRecordset.State: "&objRecordset.State
DBG "FWX.Data.OpenRS> A) IsObject(objRecordset.ActiveConnection): "&IsObject(objRecordset.ActiveConnection)
If .State=1 Then
DBG "FWX.Data.OpenRS> Close previourly open recordset"
.Close()
End If
DBG "FWX.Data.OpenRS> B) objRecordset.State: "&objRecordset.State
DBG "FWX.Data.OpenRS> B) IsObject(objRecordset.ActiveConnection): "&IsObject(objRecordset.ActiveConnection)
DBG "FWX.Data.OpenRS> ValidConn(objRecordset)..... (TAKE 1)"
bNeedsConnection=CBool(Not ValidConn(objRecordset))
DBG "FWX.Data.OpenRS> ValidConn(objRecordset) --> bNeedsConnection: "&bNeedsConnection
If bNeedsConnection And IsObject(FWX)Then
DBG "FWX.Data.OpenRS> TEST 2: a) FWX.Setting('ConnecionString')  = "&FWX.ConnectionString
If FWX.ConnectionString="" Then
DBG "EMERGENCY REBOOT REQUIRED (ConnectionString Len("&FWX.ConnectionString&"))"
DBG.Warning "Emergency reboot required"
FWX.Reboot()
DBG "EMERGENCY REBOOT PERFORMED (ConnectionString Len("&FWX.ConnectionString&"))"
DBG "FWX.Data.OpenRS> TEST 2: b) FWX.Setting('ConnecionString') = "&FWX.ConnectionString
End If
sConn=FWX.ConnectionString
DBG "FWX.Data.OpenRS> TEST 3: a) FWX.Setting('ConnecionString') = "&FWX.ConnectionString
If sConn<>"" Then
ConnectRS objRecordset,sConn
DBG "FWX.Data.OpenRS> TEST 3: b) ValidConn(objRecordset)..... (TAKE 2)"
Err.Clear()
bNeedsConnection=Not ValidConn(objRecordset)
DBG "FWX.Data.OpenRS> TEST 3: c) bNeedsConnection = "&bNeedsConnection
End If
End If
If bNeedsConnection Then
DBG "FWX.Data.OpenRS> TEST 4: NEEDS CONN!"
sConn=Application.Contents("ConnectionString")
If IsTest Then sConn=TOT(Application.Contents("ConnectionString-Test"),sConn)
If IsLocal Then sConn=TOT(Application.Contents("ConnectionString-Local"),sConn)
DBG "FWX.Data.OpenRS> TEST 4: a) sConn = "&sConn
If sConn<>"" Then
DBG "FWX.Data.OpenRS> TEST 4: c)	ConnectRS objRecordset, sConn"
ConnectRS objRecordset,sConn
bNeedsConnection=Not ValidConn(objRecordset)
DBG "FWX.Data.OpenRS> TEST 4: d)bNeedsConnection = "&bNeedsConnection
End If
End If
If bNeedsConnection Then
SendEmail "dev@"&MyDomain,"diego.alto@gmail.com","Dead Database Connection","IN USE: "&Application.Contents("ConnectionString")&"<br/><br/><br/><br/>LIVE: "&Application.Contents("ConnectionString-Live")&"<br/>TEST: "&Application.Contents("ConnectionString-Test")&"<br/>LOCAL: "&Application.Contents("ConnectionString-Local")&"<br/>"
DBG.FatalError "Cannot reach database"'.(FWX="& Len(FWX.ConnectionString) &",App="& Len(Application.Contents("ConnectionString")) &"). Can't run sRunSQL>"&sRunSQL&""
End If
j("_STAT-LastRS")=sRunSQL
j("_STAT-OpenRS")=VOZ(j("_STAT-OpenRS"))+1
DBG "DATA> OpenRS: "&j("_STAT-OpenRS")&": "&sRunSQL
DBG "FWX.Data.OpenRS.Query=GO..."
Dim q_time,q_had_errors,q_has_errors,q_had_details,q_has_details
q_had_errors=CBool(Err.Number<>0)
q_had_details=Err.Description
Err.Clear()
DBG ""
q_time=Timer()
.Open sRunSQL
q_time=Timer()-q_time
q_has_details=Err.Description
q_has_errors=CBool(Err.Number<>0 Or q_has_details<>"")
DBG ""
If q_has_errors Then
DBG "FWX.Data.OpenRS.Query=ERROR! ("&q_has_details&")"
DBG "DATA> OpenRS: There were errors in this query"
If IsTest Or IsBack Then
KillRS objRecordset
DBG.FatalError "/*SQL ERROR: "&TOT(q_has_details,"Unknown")&Enclose(q_had_details," (PREVIOUS ERROR:",")")&" */<br/>"&vbCRLf&sRunSQL
End If
Else
DBG "FWX.Data.OpenRS.Query=NO ERRORS!"
End If
If IsTest Or IsStaffPC Or DBG.Show Then
j("_STAT-OpenRS-List")=PutAfter(j("_STAT-OpenRS-List"),vbCRLf)'&"/*-*/"&vbCRLf)
j("_STAT-OpenRS-List")=j("_STAT-OpenRS-List")&"/*Q"&Digits(j("_STAT-OpenRS"),2)&"*/"&"/*T"&Left(Replace(FormatNumber(q_time,6,0),",",""),6)&"ms*/"&Encode(sRunSQL)&""
j("_STAT-OpenRS-List")=Replace(j("_STAT-OpenRS-List"),"[/**/","[")
j("_STAT-OpenRS-List")=Replace(j("_STAT-OpenRS-List"),"[","[/**/")
End If
If Err.Description<>"" Or Err.Number<>0 Then
Dim sError:sError=""
sError=sError&sRunSQL&"<br/>"&vbCRLf
sError=sError&Err.Description&"<br/>"&vbCRLf
DBG.FatalError sError
End If
DBG "FWX.Data.OpenRS.Query=COMPLETE!"
If .State=1 Then
DBG "FWX.Data.OpenRS.RecordCount = "&.RecordCount
Else
DBG "FWX.Data.OpenRS.RecordCount = 0 (found nothing)"
End If
If bNeedsConnection Then
Set .ActiveConnection=Nothing
DBG "DATA> OpenRS: ### KILLED TEMP CONNECTION! ###"
DBG "FWX.Data.OpenRS.END - ### KILLED TEMP CONNECTION! ###"
End If
End With
DBG "FWX.Data.OpenRS.END - SQL: "&sRunSQL&""
DBG "FWX.Data.OpenRS.END - sConn: "&sConn&""
DBG "FWX.Data.OpenRS.END - Query: http://"&SERVER_NAME&FULLURL&""
DBG "FWX.Data.OpenRS.END - server: "&SERVER_NAME&""
DBG "FWX.Data.OpenRS.END - MyDomain: "&MyDomain&""
OpenRS=Exists(objRecordset)
End Function
Set j("XmlFiles")=Dictionary()
Function OpenXML(ByVal lcl_sFileName)
Err.Clear():On Error Resume Next
If IsEmpty(lcl_sFileName)Then Exit Function
j("OpenXml")=j("OpenXml")+1
j("XmlFiles")("F"&j("OpenXml")&"")=lcl_sFileName
lcl_sFileName=PhysicalPath(lcl_sFileName)
DBG "XML: opening xml file '"&lcl_sFileName&"'"
Set OpenXML=RecordSet()
OpenXML.Open lcl_sFileName,"Provider=MSPersist;",adOpenForwardOnly,adLockReadOnly,adCmdFile
If Err.Description<>"" Or Err.Number<>0 Then
DBG "XML: XML-OPEN-FAILED! '"&lcl_sFileName&"'"
Err.Clear()
Set OpenXML=RecordSet()
OpenXML.Open()
Else
DBG "XML: XML-OPEN-OK '"&lcl_sFileName&"' ("&OpenXML.RecordCount&")"
End If
End Function
Function OpenUrl(ByVal s)
Set OpenUrl=Server.CreateObject("ADODB.Recordset")
OpenUrl.Open s&"",,adOpenStatic,adLockBatchOptimistic
If Err.Description<>"" Or Err.Number<>0 Then
DBG "ERR OpenUrl: "&s&""
Response.End()
End If
End Function
Function RecordCount(ByRef lclRecordSet)
If Not IsObject(lclRecordSet)Then Exit Function
RecordCount=lclRecordSet.RecordCount
If Err.Description<>"" Or Err.Number<>0 Then RecordCount=-1
Err.Clear()
End Function
Function RSSource(ByRef lclRecordSet)
If Not IsObject(lclRecordSet)Then Exit Function
RSSource=lclRecordSet.Source
If Err.Description<>"" Or Err.Number<>0 Then RSSource=""
Err.Clear()
End Function
Function SaveFromRS(ByRef source,ByRef dest,ByVal prefix)
Err.Clear():On Error Resume Next
If Not IsObject(source)Or Not IsObject(dest)Then Exit Function
For Each rfld In source.Fields
dest(prefix&rfld.Name)=source(rfld.Name)
Next
End Function
Function SaveAllFromRS(ByRef source,ByRef dest,ByVal key)
If Not IsObject(source)Or Not IsObject(dest)Then Exit Function
Do While Exists(source)
dest.Filter="["&key&"] = "&source(key)
If Not Exists(dest)Then dest.AddNew
SaveFromRS source,dest,prefix
source.MoveNext
Loop
End Function
Function SaveToRS(ByRef source,ByRef dest)
SaveToRS=DoSaveToRS(source,dest,"","")
End Function
Function SaveToRSWithSufix(ByRef source,ByRef dest,ByVal sufix)
SaveToRSWithSufix=DoSaveToRS(source,dest,"",sufix)
End Function
Function SaveToRSWithPrefix(ByRef source,ByRef dest,ByVal prefix)
SaveToRSWithPrefix=DoSaveToRS(source,dest,prefix,"")
End Function
Function DoSaveToRS(ByRef source,ByRef dest,ByVal prefix,ByVal sufix)
If Not IsObject(source)Or Not IsObject(dest)Then Exit Function
For Each fld In dest.Fields
If source(prefix&fld.Name&sufix)="{Current}" Then
Else
If 1=0 Then
ElseIf fld.Type=11 Then
dest(fld.Name)=IsTrue(source(prefix&fld.Name&sufix))
RSFieldSet dest,fld.Name,IsTrue(source(prefix&fld.Name&sufix))
ElseIf IsSet(source(prefix&fld.Name&sufix))Then
If fld.Name="date" Or fld.Type=135 Then
RSFieldSet dest,fld.Name,CDate(source(prefix&fld.Name&sufix))
Else
RSFieldSet dest,fld.Name,source(prefix&fld.Name&sufix)
End If
Else
RSFieldSet dest,fld.Name,source(prefix&fld.Name&sufix)
End If
End If
Next
End Function
Function SaveAllToRS(ByRef source,ByRef dest,ByVal key)
Do While Exists(source)
dest.Filter="["&key&"] = "&source(key)
If Not Exists(dest)Then dest.AddNew
SaveToRS source,dest
source.MoveNext
Loop
End Function
Function SaveXML(ByRef rsSource,ByVal lcl_sFileName)
Err.Clear():On Error Resume Next
If IsEmpty(lcl_sFileName)Then Exit Function
j("SaveXml")=j("SaveXml")+1
lcl_sFolder=FolderPath(lcl_sFileName)
If Not FolderExists(lcl_sFolder)Then BuildPath lcl_sFolder
lcl_sFileName=FixPath(lcl_sFileName)
lcl_sFileName=PhysicalPath(lcl_sFileName)
DBG "DELETE XML FILE: "&lcl_sFileName&""
DeleteFile lcl_sFileName
Err.Clear()
DBG "SAVE XML FILE: "&lcl_sFileName&""
rsSource.Save lcl_sFileName,1
If Err.Description<>"" Or Err.Number<>0 Then
DBG "XML-SAVE-FAILED!<br/><br/><p>'"&Err.Description&"'</p><br/><br/>"&lcl_sFileName
Else
DBG "XML: XML-SAVE-OK! '"&lcl_sFileName&"'"
End If
Err.Clear()
End Function
Dim LAST_TEST_CONNECTION:LAST_TEST_CONNECTION=""
Dim TESTED_CONNECTIONS:Set TESTED_CONNECTIONS=Dictionary()
Function TestConnectionString__OPEN(ByRef cT,ByVal conn)
On Error Resume Next
Err.Clear()
cT.Open conn
If Err.Number=0 Then TestConnectionString__OPEN=True
Err.Clear()
End Function
Function TestConnectionString(ByVal conn)
Dim bT,m5
If TOE(conn)="" Then Exit Function
m5=MD5(conn)
If Application.Contents(m5)<>"" Then
bT=IsTrue(Application.Contents(m5))
Else
bT=False
Dim cT
Set cT=Server.CreateObject("ADODB.Connection")
bT=TestConnectionString__OPEN(cT,conn)
TestConnectionString__CLOSE cT
KillRS cT
Err.Clear()
If bT Then
Application.Contents(m5)="y"
End If
End If
TestConnectionString=bT
End Function
Function TestConnectionString__CLOSE(ByRef conn)
On Error Resume Next
Err.Clear()
cT.Close()
Err.Clear()
End Function
Function TestConnectionObject(ByRef conn)
On Error Resume Next
Err.Clear()
Dim bT
bT=False
If conn.State=1 Then bT=True
If Err.Number<>0 Then bT=False
Err.Clear()
TestConnectionObject=bT
End Function
Function TestConnection(ByRef conn)
Dim bT
bT=False
If IsObject(conn)Then
bT=TestConnectionObject(conn)
Else
bT=TestConnectionString(conn)
End If
TestConnection=bT
End Function
Function TestSQL(ByRef objRecordset,ByVal lcls_SQL)
Err.Clear():On Error Resume Next
Dim blnResult:blnResult=False
If Not IsObject(objRecordset)Then CreateRS objRecordset
blnTempConn=False
With objRecordset
Dim blnTempConn:blnTempConn=Not ValidConn(objRecordset)
If blnTempConn Then
ConnectRS objRecordset,FWX.ConnectionString
Else
End If
If .State=1 Then .Close()
.Open lcls_SQL
If Err.number=0 Then
blnResult=True
Else
End If
If blnTempConn Then Set .ActiveConnection=Nothing
End With
TestSQL=blnResult
End Function
Function UpdateRS(ByRef objRecordset)
blnTempConn=Not ValidConn(objRecordset)
If blnTempConn Then
ConnectRS objRecordset,FWX.ConnectionString
If Err.Description<>"" Or Err.Number<>0 Then
DIE "DATA:CONNECTION FAILED"
End If
End If
Err.Clear()
objRecordset.Update()
If Err.Description<>"" Or Err.Number<>0 Then
Dim s:Set s=New QuickString:s.Divider="<br/>"&vbCRLf
s.Add "UpdateRS Error: "&Err.Description
If IsObject(objRecordset)Then
s.Add "Redordset: "&DescribeRS(objRecordset)
If IsObject(objRecordset.ActiveCommand)Then
s.Add "Command: "&objRecordset.ActiveCommand.CommandText
End If
End If
DBG.FatalError s
End If
If blnTempConn Then
Set objRecordset.ActiveConnection=Nothing
End If
End Function
Function ValidConn(ByRef objRecordset)
Dim blnResult
blnResult=False
If IsObject(objRecordset)Then
blnResult=TestConnection(objRecordset.ActiveConnection)
Else
blnResult=TestConnection(objRecordset)
End If
ValidConn=blnResult
End Function
Function Rs2Dic(ByVal source):Set Rs2Dic=RsToDic(source):End Function
Function RsToDic(ByVal source)
Err.Clear():On Error Resume Next
Dim output
Set output=Dictionary()
Dim oFlds
Set oFlds=source.Fields
If Err.Description<>"" Or Err.Number<>0 Then
For Each f In source
If Not IsObject(source(f))Then
output(f)=source(f)
End If
Next
Else
For Each oFld In oFlds
output(oFld.Name)=oFld.Value
Next
End If
Set RsToDic=output
End Function
Function Dic2Rs(ByVal d):Set Dic2Rs=DicToRs(d):End Function
Function DicToRs(ByVal d)
Dim rs,s,done
Set rs=Server.CreateObject("ADODB.Recordset")
For Each s In d
s=LCase(s)
If Not rxTest(","&done&",",","&s&",")Then
done=done&","&s
rs.Fields.Append s,adVarChar,500
End If
Next
rs.Open
rs.AddNew
For Each s In d
rs.Fields(s)=d(s)
Next
rs.MoveLast
Set DicToRs=rs
End Function
Function GetArrayItem(ByVal a,ByVal i)
On Error Resume Next
GetArrayItem=a(i)
End Function
Function GetFieldType(ByRef rs,ByVal f)
Dim ft
ft=RSFieldType(GetFieldObj(rs,f))
GetFieldType=ft
End Function
Function GetField(ByRef rs,ByVal f)
Err.Clear():On Error Resume Next
Dim output
output=rs(f)
GetField=output
Err.Clear()
End Function
Function GetFieldObj(ByRef rs,ByVal f)
Err.Clear():On Error Resume Next
Dim output
Set output=rs(f)
If Err.Number<>0 Then
Set output=Nothing
End If
Set GetFieldObj=output
Err.Clear()
End Function
Const SQLKeywords="order|in|not|save|where|date|last|from|to|sum|primary|file|by|group"
Function SQLSafeColumn(ByVal s)
SQLSafeColumn=SQLSafeField(s)
End Function
Function SQLSafeField(ByVal s)
s=Trim(s)
s=rxReplace(s,"^(\w\.)?("&SQLKeywords&")$","$1[$2]")
SQLSafeField=s
End Function
Function SQLSafeString(ByVal s)
s=Replace(TOE(s),"'","''")
Dim bad,x,y
bad=Array("[","]")
For Each x In bad
If InStr(1,s,x)Then
s=Replace(s,x,"_C1_"&ASC(Left(x,1))&"_C2_"&Mid(x,2))
End If
Next
s=Replace(s,"_C1_","'+CHAR(")
s=Replace(s,"_C2_",")+'")
SQLSafeString=s
End Function
Function SQLSafeBoolean(ByVal s)
Select Case LCase(s&"")
Case "0","n","f","false","no","not","":
s=0
Case Else:
s=1
End Select
SQLSafeBoolean=s
End Function
Function SQLSafeDate(ByVal s)
s=CDate(s)
SQLSafeDate=s
End Function
Function SQLSafeNumber(ByVal n)
Err.Clear():On Error Resume Next
n=CDbl(n)
If Err.Description<>"" Or Err.Number<>0 Then n=0:Err.Clear()
SQLSafeNumber=n
End Function
Function RSPaging(ByRef rsNav,ByVal strOptions)
Dim output:Set output=New QuickString
Dim iPage,iPags,iSize,iRecs,iUpTo,bAuto,sNext,bGradual
bAuto=IsSet(Query.FQ("RSPaging-auto"))
iPage=VOZ(VOV(Query.FQ("@page"),1))
iPags=VOZ(rsNav.PageCount)
iSize=VOZ(rsNav.PageSize)
iRecs=VOZ(rsNav.RecordCount)
iUpTo=iPage*iSize
sNext=rxReplace(ReSubmit("RSPaging=y&~RSPaging-auto&@size="&iSize&"&@page="&iPage+1),"a=(lis[^&]+|top[^&]+)&","a=list-i&")
sCount=Replace(ReSubmit("~RSPaging&~RSPaging-auto&~@size&~@page&a=tally&"),"a=list&","a=tally&")
bGradual=CBool(InStr(1,strOptions,"gradual")>0)
If Not(iPags>1)Then Exit Function
If IsTrue(j("-PRINT"))Then
output.Add "<div class='RSPaging-print Clear' style=""padding:3px;"">"&vbCRLf
output.Add " Page <b>"&iPage&"</b> of <b>"&iPags&"</b>"&vbCRLf
output.Add "</div>"&vbCRLf
RSPaging=output
Exit Function
End If
If iPags>1 Then
If iPage=iPags Then
output.Add "<scr"&"ipt>$('#RSPaging-auto-notice,.RSPaging-auto-notice').remove();</scr"&"ipt>"
Else
If Query.FQ("RSPaging-auto")="" Or VOZ(Query.FQ("@page"))=0 Then
output.Add "<div id='RSPaging-auto-notice' class='RSPaging-auto-notice' style='left:50%; margin:-10px auto 0 -45px;  padding:10px 10px 3px 10px;  display:"&IIf(bAuto,"block","none")&";position:fixed;top:0;  border-radius:10px;  width:180px;z-index:99999;background:red;color:white;font-weight:bold;border:red solid 1px;font-size:24px;text-align:center;'>"
output.Add "Loading..."
output.Add "</div>"
Else
output.Add "<scr"&"ipt>$('#RSPaging-auto-notice').show();</scr"&"ipt>"
End If
End If
End If
If iPage<iPags Then
output.Add "<div id='RSPaging' class='RSPaging pHide {"
output.Add "max:"&iRecs&","
output.Add "page:"&iPage&","
output.Add "size:"&iSize&","
output.Add "recs:"&iRecs&","
output.Add "upto:"&iUpTo&","
output.Add "pages:"&iPags&","
output.Add "pages:"&iPags&","
output.Add "auto:"&IIf(bAuto,"true","false")&""
output.Add "}'>"
output.Add "<div class='RSPaging-print pShow'>"&vbCRLf
output.Add "Showing the first <span class='B'>"&iUpTo&" items"&IIf(bGradual,""," of a total "&iRecs&"")&"</span>"
output.Add "</div>"&vbCRLf
output.Add "<div class='RSPaging-links Clear pHide' style='text-align:right;'>"&vbCRLf
output.Add "<a id='RSPaging-next' class='button auto-appear'"
output.Add " onclick=""return AjaxReplace($(this).parents('.RSPaging'),this.href);"""
output.Add " href='"&sNext&"'"
output.Add " style='float:left;'"
output.Add ">"
output.Add "Showing <span class='B'>"&iUpTo&" of "&IIf(bGradual,"<span class='many'>many</span>",iRecs)&"</span>: "
If bGradual Then
output.Add "<span class='x-more-vague"&IIf(bGradual,""," x-Hidden")&"'>show more...</span>"
Else
If(iUpTo+iSize)>iRecs Then
output.Add "<span class='more-known"&IIf(bGradual," x-Hidden","")&"'>show "&(iRecs-iUpTo)&" more...</span>"
Else
output.Add "<span class='more-known"&IIf(bGradual," x-Hidden","")&"'>show "&iSize&" more...</span>"
End If
End If
output.Add "</a>"&vbCRLf
If bGradual And IsBack And StartsWith(sCount,"/x/@")Then
output.Add "<a id='RSPaging-count' class='x-button auto {delay:"&IIf(IsTest,3,10)&"} RSPaging-count'"
output.Add " href='"&sCount&"'"
output.Add ">"
output.Add "<span class='count'>how many?</span>"
output.Add "</a>"&vbCRLf
End If
output.Add "<a id='RSPaging-auto' class='button hFade"&IIf(bAuto," auto","")&"'"
output.Add " title='Show all "&IIf(bGradual,"",(iRecs-iSize)&" ")&"remaining records'"
output.Add " onclick=""$('#RSPaging-auto-notice').show(); return AjaxReplace($(this).parents('.RSPaging'),this.href);"""
output.Add " href='"&sNext&"RSPaging-auto=y&'"
output.Add ">"
output.Add "<span>show all</span>"
output.Add "</a>"&vbCRLf
output.Add "</div>"&vbCRLf
output.Add "</div>"&vbCRLf
End If
RSPaging=output
End Function
Function RSFieldType(ByRef fo)
Dim ft
If IsObject(fo)Then
If Not fo Is Nothing Then
If 1=0 Then
ElseIf RSFieldBoolean(fo)Then:ft="boolean"
ElseIf RSFieldNumeric(fo)Then:ft="number"
ElseIf RSFieldDate(fo)Then:ft="date"
ElseIf RSFieldText(fo)Then:ft="text"
ElseIf RSFieldBlob(fo)Then:ft="blob"
Else
ft=""
End If
End If
End If
RSFieldType=ft
End Function
Function RSFieldBoolean(ByRef fld)
RSFieldBoolean=CBool(fld.Type=11)
End Function
Function RSFieldDate(ByRef fld)
RSFieldDate=CBool(fld.Type=7 Or fld.Type=135)
End Function
Function RSFieldBlob(ByRef fld)
RSFieldBlob=CBool(fld.Type=128 Or fld.Type=203)
End Function
Function RSFieldNumeric(ByRef fld)
RSFieldNumeric=CBool(fld.Type<=6 Or fld.Type=14 Or fld.Type=17 Or fld.Type=20 Or fld.Type=131)
End Function
Function RSFieldText(ByRef fld)
RSFieldText=Not(RSFieldNumeric(fld)Or RSFieldDate(fld)Or RSFieldBoolean(fld))
End Function
Function RSCount(ByRef rD)
Err.Clear():On Error Resume Next
If Exists(rD)Then
RSCount=rD.RecordCount
End If
Err.Clear()
End Function
Function RSFieldGet(ByRef rD,ByVal f)
Err.Clear():On Error Resume Next
o=rD(f)
If Err.Description="" And Err.Number=0 Then
RSFieldGet=o
Else
Err.Clear()
End If
End Function
Function RSFieldSet(ByRef rD,ByVal f,ByVal v)
RSFieldSet=RSFieldLet(rD,f,v)
End Function
Function RSFieldLet(ByRef rD,ByVal f,ByVal v)
Err.Clear():On Error Resume Next
Dim o:Set o=rD(f)
If Err.Description="" And Err.Number=0 Then
If RSFieldBoolean(o)Then
o.Value=IsTrue(v)
ElseIf RSFieldDate(o)Then
v=MakeDateFromString(v)
If v<>"" Then
o.Value=IIf(v<>"",v,vbNull)
End If
ElseIf RSFieldNumeric(o)Then
o.Value=VOZ(v)
Else
v=ASCII(v)
If Len(v)>o.DefinedSize Then
v=Left(v,o.DefinedSize)
End If
o.Value=v
End If
End If
RSFieldLet=CBool(Err.Number=0)
Err.Clear()
End Function
Function RSFieldObj(ByRef rD,ByVal f)
Err.Clear():On Error Resume Next
Set o=rD(f)
If Err.Description="" And Err.Number=0 Then
Set RSFieldObj=o
Else
Set RSFieldObj=Nothing
Err.Clear()
End If
End Function
Function IsRSField(ByRef rD,ByVal f)
Set o=RSFieldObj(rD,f)
If Err.Description="" And Err.Number=0 Then
If o Is Nothing Then
IsRSField=False
Else
IsRSField=True
End If
Else
IsRSField=False
Err.Clear()
End If
Set o=Nothing
End Function
Dim RSKiller
Set RSKiller=New FWX_AutomaticRecordsetKiller
Class FWX_AutomaticRecordsetKiller
Public AKR_ARRAY()
Public Length
Private Sub Class_Initialize()
Me.Initialize()
End Sub
Private Sub Class_Terminate()
Me.Terminate()
End Sub
Public Sub Initialize()
Me.Length=0
End Sub
Public Sub Register(ByRef rs)
Me.Length=Me.Length+1
ReDim Preserve AKR_ARRAY(Me.Length)
Set AKR_ARRAY(Me.Length-1)=rs
DBG "RSKiller> REGISTER > Me.Length = "&Me.Length
End Sub
Public Function IsOpen(ByRef rs)
On Error Resume Next
If rs.State=1 Then
IsOpen=True
End If
If Err.Number<>0 Then
Err.Clear()
End If
End Function
Public Sub Kill(ByRef rs)
On Error Resume Next
rs.Close()
Err.Clear()
Set rs=Nothing
End Sub
Public Sub Terminate()
DBG "RSKiller> TERMINATE (kill all "&Me.Length&" recordsets)"
Dim x,z,leaks
For x=1 To Me.Length
z=x-1
If IsObject(Me.AKR_ARRAY(z))Then
DBG "RSKiller> TERMINATE > RS"&x
DBG "RSKiller> TERMINATE > RS"&x&" > "&RSSource(Me.AKR_ARRAY(z))
If Me.IsOpen(Me.AKR_ARRAY(z))Then
leaks=leaks+1
DBG "RSKiller> TERMINATE > RS"&x&" > IS A LEAK!"
If IsTest And DBG.Show Then DBG PrintRS(Me.AKR_ARRAY(z))
Me.Kill(Me.AKR_ARRAY(z))
DBG "RSKiller> TERMINATE > RS"&x&" > KILLED!"
End If
End If
Next
DBG "RSKiller> TERMINATE > ALL DONE!"
DBG "RSKiller> TERMINATE > ALL DONE! > FOUND "&VOZ(leaks)&" LEAKS!"
End Sub
End Class
Dim parser_live_done
parser_live_done=False
Dim parser:Set parser=New fwxParser
Class fwxParser
Function Merge(ByVal s)
Merge=Me.DoMerge(s,True,True)
End Function
Function PreMerge(ByVal s)
PreMerge=Me.DoMerge(s,False,True)
End Function
Function DoMerge(ByVal s,ByVal bParseLIVE,ByVal bParseSEMI)
s=Me.MMM(s,bParseLIVE,bParseSEMI)
DoMerge=s
End Function
Function MMM(ByVal s,ByVal bParseLIVE,ByVal bParseSEMI)
s=TOE(s)
If s="" Then Exit Function
If(bParseLIVE)Then
s=Me.MergeGroup(bParseLIVE,s,"!","visitor,Session.Contents,Request.Cookies")',MySession,MyCookies,quote.Account")
If IsObject(Query)Then
s=Me.MergeGroup(bParseLIVE,s,"?","Query.Fata,Query.Data")
End If
End If
If(bParseLIVE Or bParseSEMI)And IsObject(j)Then
s=Me.MergeGroup(bParseLIVE,s,"#","j")
End If
s=Me.MergeGroup(bParseLIVE,s,"$","SITE,APP")',Application.Contents")
MMM=s
End Function
Function MergeGroup(ByVal bParseLIVE,ByVal s,ByVal sDelimiter,ByVal entities)
Dim sText,sTags
Dim iLevel
Dim oTags,oTag
Dim ents,ent,col
Dim fld,opt,rpc
sText=s
sAllowLive=""
If bParseLIVE Then sAllowLive="\*?"
sTags="\"&sDelimiter&sAllowLive&Replace(Me.MergeText,"\"&sDelimiter&"","")&"\"&sDelimiter&""
Set oTags=rxMatches(sText,sTags)
iLevel=0
Do While oTags.Count>0 And iLevel<3
iLevel=iLevel+1
For Each oTag In oTags
rpc=""
fld=""
fld=Mid(oTag.Value&"",2,Len(oTag.Value)-2)
If Left(fld,1)="*" And bParseLIVE Then
fld=Mid(fld,2)
End If
opt=rxSelect(fld,"\:.+$")
If opt<>"" Then
fld=Replace(fld,opt,"")
opt=Replace(opt,":","")
End If
ents=rxReplace(rxSelect(fld,"^([\.\w]+)~"),"~$","")
If ents="" Then ents=entities
If rxTest(fld,"^(SITE|SHOP|INVY|EMAN)\.")Then
ents=rxReplace(ents,"^((SITE|SHOP|APP|INVY|EMAN),)+","APP,")
End If
Dim exception:exception=False
Dim exceptional:exceptional=False
For Each ent In Split(ents,",")
If rpc="" And exceptional=False Then
fld=Mid(fld,InStr(1,fld,"~")+1)
exception=False
If sDelimiter="!" Then
If ent="visitor" Then
If rxTest(fld,"[^\w\-\_]")Then
exception=True
End If
End If
End If
If sDelimiter="$" Then
If Left(fld,1)="#" Then
exception=True
exceptional=True
rpc=FWX.Setting(fld)
ElseIf rxTest(fld,"(datasource|datasrc|database|connect|password|pwd)")And(IsFront Or(Not IsStaff))Then
exception=True
exceptional=True
ElseIf rxTest(fld,"^(jssor)")Then
exception=True
exceptional=True
rpc="$"&fld&"$"
Else
End If
End If
If ent="access" Then
exception=True
exceptional=True
Select Case LCase(fld)
Case "credentials"
rpc=MyCredentials
Case "credential"
rpc=rxSelect(MyCredentials,"^\w+")
Case "username"
rpc=rxSelect(MyEmail,"^\w+")
Case "account"
rpc=MyAccount
Case "email"
rpc=MyEmail
Case "domain"
rpc=rxSelect(MyEmail,"[^@]+$")
Case "name"
rpc=MyName
Case "forename"
rpc=rxSelect(MyName,"^\S+")
Case "surname"
rpc=Trim(rxSelect(MyName,"\s\S+$"))
Case Else:
rpc=HasAccess(MyCredentials,fld)
End Select
End If
If ent="cani" Then
exception=True
exceptional=True
End If
If exception Or exceptional Then
Else
Set col=OBJ2(ent&"")
If Not col Is Nothing Then
rpc=MergeTemplateField(col,fld,FALSE)
Else
End If
Set col=Nothing
End If
End If
Next
If opt<>"" Then
rpc=parser.Format(rpc,opt)
End If
If InStr(1,rpc,"#"&fld&"#")>0 Then
rpc=Replace(rpc,"#"&fld&"#","")
End If
s=Replace(s&"",oTag.Value&"",rpc&"")
Next
sText=s
Set oTags=rxMatches(sText,sTags)
Loop
MergeGroup=s
End Function
Function Filter(ByVal s)
Filter=Me.DoFilter(s,True,True)
End Function
Function PreFilter(ByVal s)
PreFilter=Me.DoFilter(s,False,True)
End Function
Function DoFilter(ByVal s,bParseLIVE,bParseSEMI)
If rxTest(s,"\{\/?[\w\.]+\}")Then
If bParseLIVE Then
s=Me.FilterYY(s,IsLoggedIn,"Members?|Clients?|((Logged|Signed)(In|On)?)")
s=Me.FilterYY(s,IsAjax,"Ajax|HTTPReq|HTTPRequest")
s=Me.FilterYY(s,IsAjax And Query.FQ("RSPaging")<>"","(Ajax)?(RS)?Paging|AjaxRSPaging")
s=Me.FilterYN(s,IsIE,"IE|IE6|IE7|IE8|IE9","FF|Mozil?la|Safari|Opera")
s=Me.FilterYY(s,IsMobile,"Mobile|Mob|M")
s=Me.FilterYY(s,IsMobile Or IsMobileDevice,"MobileDevice|Portable|Handheld|Android|iPhone")
s=Me.FilterYY(s,IsRobot,"Bot|Robot|Robots|Spider|Bots|SE|DotBot")
s=Me.FilterYY(s,IsHuman,"Man|Human|Person|People")
s=Me.FilterYY(s,IsMaster,"Master")
s=Me.FilterYY(s,IsAdmin,"Admin")
s=Me.FilterYY(s,IsManager,"Manager")
s=Me.FilterYY(s,IsAccountant,"Accountant|Accounting")
s=Me.FilterYY(s,IsSupervisor,"Supervisor")
s=Me.FilterYY(s,IsMarketing,"Marketing")
s=Me.FilterYY(s,IsEditor,"Editor")
s=Me.FilterYY(s,IsController,"Controller")
s=Me.FilterYY(s,IsStock,"(Stock|Warehouse|Dispatch)(Manager|Staff)?")
s=Me.FilterYY(s,IsClerk,"Clerk|Minion")
s=Me.FilterYY(s,IsStaff,"Staff")
If IsObject(DBG)Then
s=Me.FilterYY(s,DBG.Show,"Debug|Debugging")
End If
s=Me.FilterYY(s,IsStaffPC,"Staff(PC|Machine|Workstation|Computer)")
s=Me.FilterYN(s,CBool(MyState<>""),"Visitor|Initiated|Started","Nobody|Anonymous|Fresh|Stateless")
End If
If bParseSEMI Then
If j("-PDF")="" Then j("-PDF")=rxTest(Query.EXT,"^(pdf)$")
If j("-FEED")="" Then j("-FEED")=rxTest(Query.EXT,"^(xml|csv|tsv|txt)$")
s=Me.FilterYY(s,IsTemp,"Temp")
s=Me.FilterYY(s,IsTemp Or IsTest Or IsAdmin Or IsLocal,"Test|Tester|Dev")
s=Me.FilterYN(s,IsLive,"Live|Online","Local")
s=Me.FilterYY(s,j("-FEED"),"rss|feed|xml|export|doc")
s=Me.FilterYY(s,j("-PDF"),"PDF")
s=Me.FilterYY(s,IsSet(j("q")),"Search|Q")
If IsFront Then
s=Me.FilterYN(s,j("-HC"),"HC","Graphics?")
s=Me.FilterYN(s,j("-TEXT"),"Text","Graphics?")
s=Me.FilterYN(s,j("-PRINT"),"Print(er)?(Friendly)?","Screen")
s=Me.FilterYY(s,j("IsMini"),"compact|light|lite|popup|inset|mini")
s=Me.FilterYY(s,j("-INTERFACE"),"Interface")
s=Me.FilterYY(s,j("-INTERNAL")Or j("-PAGE"),"Internal|InternalPage")
s=Me.FilterYY(s,j("-ACTION"),"Act(ion)?")
s=Me.FilterYY(s,j("-FUNCTION"),"Func(tion)?")
s=Me.FilterYY(s,j("-ROOT"),"(On)?(Root|Base)")
s=Me.FilterYN(s,GTZ(j("page")),"Content","Home")
s=Me.FilterYN(s,CBool(j("k")="page" And GTZ(j("page"))),"Page|Parent","Product|Event|Post|Child")
For Each m In Split("Cart,Client,Do,Page,Product,Event",",")
s=Me.FilterYY(s,CBool(j("m")=UCase(m)),m&"(Page|Module|Mod)")
Next
If j("-FP")="" Then j("-FP")=CBool(Query.RAWURL="/index.asp" And Query.QUERY="")
s=Me.FilterYY(s,j("-FP"),"Front|FP")
If IsObject(site)Then
If InStr(1,s,"adwords}")>0 Then s=Me.FilterYY(s,IsSet(site.CFG("google.adwords.account")),"Adwords")
If InStr(1,s,"raygun}")>0 Then s=Me.FilterYY(s,IsSet(site.CFG("raygun.api")),"Adwords")
End If
Dim a
For Each a In Split("f,p",",")
x=rxReplace(j(a&"."),"\.(\w+)","(\.?$1)?")
If x<>"" Then
s=Me.FilterYY(s,True,x)
Else
If a="p" Then
s=Me.FilterYY(s,False,"PAGES(\.?\w+)?")
Else
s=Me.FilterYY(s,False,"FEEDS(\.?\w+)?")
End If
End If
Next
End If
End If
If bParseLIVE Then
s=Me.AllFiltersDone(s)
End If
End If
DoFilter=s
End Function
Function FilterYY(ByVal s,ByVal b,ByVal c)
If IsTrue(b)Then
s=FilterBlockN(s,c)
Else
s=FilterBlockY(s,c)
End If
FilterYY=s
End Function
Function FilterYN(ByVal s,ByVal b,ByVal y,ByVal n)
If IsTrue(b)Then
s=FilterBlockY(s,n)
s=FilterBlockN(s,y)
Else
s=FilterBlockY(s,y)
s=FilterBlockN(s,n)
End If
FilterYN=s
End Function
Function FilterBlockY(ByVal s,ByVal c)
s=FilterOut(s,"(Only)?(Is|For)?(In|On)?\.?"&"("&c&")"&"(\.?(Page|Mode)?(Only)?)?")
s=FilterOKd(s,"(Is)?Not(In|For)?\.?"&"("&c&")"&"(\.?(Page|Mode)?(Only)?)?")
FilterBlockY=s
End Function
Function FilterBlockN(ByVal s,ByVal c)
s=FilterOut(s,"(Is)?Not(In|For)?\.?"&"("&c&")"&"(\.?(Page|Mode)?(Only)?)?")
s=FilterOKd(s,"(Only)?(Is|For)?(In|On)?\.?"&"("&c&")"&"(\.?(Page|Mode)?(Only)?)?")
FilterBlockN=s
End Function
Function FilterOut(ByVal s,ByVal tagName)
Dim bE:bE=False
Dim rf:Set rf=RegularExpression()
rf.Global=False
rf.Pattern="\{"&tagName&"\}"
Do While rf.Test(s)And Not bE
Set matches=rf.Execute(s)
For Each m In matches
givenName=Mid(m.Value,2,Len(m.Value)-2)
closingTag=InStr(m.FirstIndex+1,s&"","{/"&givenName&"}")
If closingTag>0 Then
tagContent=Mid(s,m.FirstIndex+1,closingTag+Len("{/"&givenName&"}")-m.FirstIndex-1)
s=Left(s,m.FirstIndex)&Right(s,Len(s)+1-(closingTag+Len("{/"&givenName&"}")))
Else
bE=True
End If
Next
Set matches=Nothing
Loop
Set rf=Nothing
FilterOut=s
End Function
Function FilterOKd(ByVal s,ByVal f)
For Each m In rxMatches(s,"\{("&f&")\}")
m=Mid(m,2,Len(m)-2)
If rxTest(s,"\{/"&m&"\}")Then
s=rxReplace(s,"\{/?("&m&")\}","")
End If
Next
FilterOKd=s
End Function
Function AllFiltersDone(ByVal s)
For Each m In rxMatches(s,"\{[\w\.]+\}")
m=Mid(m,2,Len(m)-2)
If rxTest(s,"\{/"&m&"\}")Then
s=rxReplace(s,"\{/?("&m&")\}","")
End If
Next
AllFiltersDone=s
End Function
Function ParseItems(ByVal items)
Set ParseItems=Me.ParseColItems(j,items)
End Function
Function ParseColItems(ByRef col,ByVal items)
Dim i
For Each i In Split(items,",")
If Not IsObject(col(i))Then
col(i)=Me.Parse(col(i))
End if
Next
Set ParseColItems=col
End Function
Function PreParseItems(ByVal items)
Set PreParseItems=Me.PreParseColItems(j,items)
End Function
Function PreParseColItems(ByRef col,ByVal items)
Dim i
For Each i In Split(items,",")
If Not IsObject(col(i))Then
col(i)=Me.PreParse(col(i))
End if
Next
Set PreParseColItems=col
End Function
Function Prepare(ByVal s):Prepare=PreParse(s):End Function
Function PreParse(ByVal s)
PreParse=Me.DoParse(s,False,True)
End Function
Function Process(ByVal s):Process=Me.Parse(s):End Function
Public Function Parse(ByVal s)
Parse=Me.DoParse(s,True,True)
End Function
Public Function DoParse(ByVal s,bParseLIVE,bParseSEMI)
s=Replace(TOE(s),"<!--[if ","<!--[?if ")
s=Me.PPP(s,bParseLIVE,bParseSEMI)
If bParseLIVE Then
s=rxReplace(s,"\^\*","*")
s=Me.PPP(s,bParseLIVE,bParseSEMI)
s=Replace(s,"[?","[")
s=rxReplace(s,"(<!--/?/?)?(\{|\})\!(\{|\})(/?/?-->)?","")
s=Replace(s,"(-(NoParse)-)","")
End If
DoParse=s
End Function
Function PPP(ByVal s,bParseLIVE,bParseSEMI)
s=TOE(s)
If s="" Then Exit Function
j("_STAT-"&IIf(bParseLIVE,"","Pre")&"Parse")=VOZ(j("_STAT-"&IIf(bParseLIVE,"","Pre")&"Parse"))+1
Dim iParseLevel,aFoundTags,iTagCount:iParseLevel=0
Dim sR,sT,oM
Dim bTagsBeingProcessed:bTagsBeingProcessed=True
s=Me.DoFilter(s,bParseLIVE,bParseSEMI)
s=Me.DoMerge(s,bParseLIVE,bParseSEMI)
Set aFoundTags=FindTags(s):iTagCount=aFoundTags.Count
Do While iTagCount>0 And iParseLevel<FWX___PARSER_MAX_PARSE_DEPTH And bTagsBeingProcessed
iParseLevel=iParseLevel+1
bTagsBeingProcessed=False
For Each oM In aFoundTags
sT=oM.Value:sR=""
If Left(sT,6)="[if IE" Or Left(sT,6)="[endif" Or Left(sT,1)=">" Then
sR=sT
End If
If InStr(1,sT,"cdnjs.cloudflare")>0 Then
sR=sT
End If
If sR="" Then
If Not bParseLIVE Then
If rxTest(sT,"^\[\*")Then
sR=sT
End If
If rxTest(sT,"\#\*"&Me.MergeText&"\#")Then
sR=sT
End If
If rxTest(sT,"\$\*"&Me.MergeText&"\$")Then
sR=sT
End If
If rxTest(sT,"\|"&Me.MergeText&"\|")Then
If rxTest(sT,"(template|value|label)\=[""'][^\|""']*\|"&Me.MergeText&"\|[^\|""']*[""']")Then
Else
sR=sT
End If
End If
End If
End If
If sR="" Then
If Not bParseSEMI Then
If rxTest(sT,"^\[\#")Then
sR=sT
End If
End If
If rxTest(sT,"^\[[\?\\\/]")Then
sR=sT
End If
End If
If sR="" Then
If rxTest(sT,"[\#\!\?\$]"&Me.MergeText&"[\#\!\?\$]")Then
sR=sT
End If
End If
If sR="" Then
If Not bL Then sR=Me.Tag(sT)
bTagsBeingProcessed=True
End If
s=Replace(s&"",sT&"",sR&"")
Next
s=Me.DoFilter(s,bParseLIVE,bParseSEMI)
s=Me.DoMerge(s,bParseLIVE,bParseSEMI)
Set aFoundTags=FindTags(s):iTagCount=aFoundTags.Count
If iParseLevel>FWX___PARSER_MAX_PARSE_DEPTH Then DBG.Warning "Deep parse detected!"
Loop
Set aFoundTags=Nothing
PPP=s
End Function
Function FindTags(ByVal s)
Set FindTags=rxMatches(s,"\[[\*\#]?[a-zA-Z]\w{1,}((\])|(\s{1,}(\s*(?!\])([\w\-\_\.]+)\s*\=\s*((\"""&"([^""]|\\"")*?"&"(?!\\)\"")|(\'"&"([^']|\\')*?"&"(?!\\)\'))\s*)*[^\]]*\]))")
End Function
Public Function oTag()
Set oTag=aTag(iTag-1)
End Function
Function Tag(ByVal sThisTagDefinition)
Tag=TagFor(sThisTagDefinition,"")
End Function
Function TagFor(ByVal sThisTagDefinition,ByRef oEnt)
If sThisTagDefinition&""="" Then Exit Function
Dim sThisTagOutput
sThisTagOutput=Me.TagCache(sThisTagDefinition)
If sThisTagOutput<>"" Then
Else
If Not iTag>0 Then iTag=1
If Not IsObject(aTag(iTag))Then
Set aTag(iTag)=New fwxParserTag
End If
With aTag(iTag)
iTag=iTag+1
.Start sThisTagDefinition
If IsObject(oEnt)Then Set .Source=oEnt
b=.Process()
sThisTagOutput=.Data("_output")
Me.TagCache(sThisTagDefinition)=sThisTagOutput
iTag=iTag-1
End With
End If
TagFor=sThisTagOutput
End Function
Function Escape(ByVal s)
s=Replace(s,"[","[?")
Escape=s
End Function
Public ToBeFormatted
Function Format(ByVal s,ByVal f)
Me.ToBeFormatted=s
Dim o
o=Me.ToBeFormatted
f=LCase(Trim(rxReplace(f,"\W+"," ")))
For Each f In Split(f," ")
If 1=0 Then
Else
Select Case LCase(f)
Case "vo1":o=VOV(o,1)
Case "num","voz","number":o=VOZ(o)
Case "cur":o=Cur(VOZ(o))
Case "abscur":o=Cur(ABS(VOZ(o)))
Case "ccur":o=ColourCur(VOZ(o))
Case "acc","cash","money":o=FormatNumber(VOZ(o),2)
Case "nlist","numlist","numberlist":o=TOT(NList(o),"0")
Case "round":o=Round(VOZ(o))
Case "percent","percentoutof100","percentage":o=FormatNumber((VOZ(o)/1)*100,1)
Case "percent5","percentoutof5","percentage5":o=FormatNumber((VOZ(o)/5)*100,1)
Case "percent10","percentoutof10","percentage10":o=FormatNumber((VOZ(o)/10)*100,1)
Case "nn":o=Digits(o,2)
Case "nnn":o=Digits(o,3)
Case "nnnn":o=Digits(o,4)
Case "nnnnn":o=Digits(o,5)
Case "nnnnnn":o=Digits(o,6)
Case "nnnnnnn":o=Digits(o,7)
Case "nnnnnnnn":o=Digits(o,8)
Case "iso":o=FmtDate(o,fwxDateISO)
Case "rfc":o=FmtDate(o,fwxDateRFC)
Case "date":o=FmtDate(o,fwxDateOnly)
Case "isdate":o=IIf(IsDate(DateFromString(o)),"y","n")
Case "makedate":o=DateFromString(o)
Case "m","month":o=FmtDate(o,"%m")
Case "mm":o=FmtDate(o,"%M")
Case "mmm":o=FmtDate(o,IIf(f="mmm","%b","%B"))
Case "d":o=FmtDate(o,"%d")
Case "dd":o=FmtDate(o,"%D")
Case "o":o=FmtDate(o,"%O")
Case "j":o=FmtDate(o,"%j")
Case "yy":o=FmtDate(o,"%y")
Case "yyyy","year":o=FmtDate(o,"%Y")
Case "w":o=FmtDate(o,"%w")
Case "a":o=FmtDate(o,"%A")
Case "aaa":o=FmtDate(o,"%a")
Case "h":o=FmtDate(o,"%h")
Case "hh":o=FmtDate(o,"%H")
Case "mn","min":o=FmtDate(o,"%N")
Case "s","ss":o=FmtDate(o,"%S")
Case "am","pm","ampm":o=FmtDate(o,"%P")
Case "unix":o=UnixDate(o)
Case "hhnn","time":
If rxTest(o,rxTime)Then
o=rxSelect(o,rxTimeOnly)
Else
o=FmtDate(o,fwxTime)
End If
Case "since":o=TimeSince(o)
Case "rfc":o=FmtDate(o,fwxDateRFC)
Case "iso":o=FmtDate(o,fwxDateISO)
Case "rss":o=FmtDate(o,fwxDateRSS)
Case "xml":o=FmtDate(o,fwxDateXML)
Case "rdf":o=FmtDate(o,fwxDateRDF)
Case "ics":o=FmtDate(o,fwxDateICS)
Case "cal":o=FmtDate(o,fwxDateCAL)
Case "ymd":o=FmtDate(o,"%y%M%D")
Case "yymmdd":o=FmtDate(o,"%y%M%D")
Case "yyyymmdd":o=FmtDate(o,"%Y%M%D")
Case "yymmddhh":o=FmtDate(o,"%y%M%D%H")
Case "yyyymmddhh":o=FmtDate(o,"%Y%M%D%H")
Case "yymmddhhnn":o=FmtDate(o,"%y%M%D%H%N")
Case "yyyymmddhhnn":o=FmtDate(o,"%Y%M%D%H%N")
Case "timestamp":o=FmtDate(o,"%Y%M%D-%H%N")
Case "shorttimestamp":o=FmtDate(o,"%M%D-%H%N")
Case "age-days","days","age":o=DateDiff("d",o,ZuluDate(Now()))
Case "age-minutes","minutes":o=DateDiff("n",o,ZuluDate(Now()))
Case "age-seconds","seconds":o=DateDiff("s",o,ZuluDate(Now()))
Case "md5":o=MD5(o)
Case "sha","sha256":o=SHA256(o)
Case "md5encrypt":o=Encrypt(MD5(o))
Case "encrypt":o=Encrypt(o)
Case "signed":o=SignedMessage(o)
Case "salted":o=SignedMessageWithSalt(o,Me.Data("salt"))
Case "annumber":o=ANNumber(o)
Case "anstring":o=ANString(o)
Case "pc":o=DoPostcode(o)
Case "phone":o=DoPhone(o)
Case "email":o=DoEmail(o)
Case "isemail":o=IsEmail(o)
Case "postcode":o=UCase(o)
Case "yn":o=IIf(IsTrue(o),"y","n")
Case "ny":o=IIf(IsTrue(o),"n","y")
Case "10":o=IIf(IsTrue(o),"1","0")
Case "bln":o=IIf(IsTrue(o),"true","false")
Case "yesno":o=IIf(IsTrue(o),"yes","no")
Case "jsbln","jsyn":o=IIf(IsTrue(o),"true","false")
Case "boolean":o=IIf(IsTrue(o),"true","false")
Case "e","doe":o=IIf(o="" Or o="-","true","false")
Case "d","dash":o=IIf(o="-","true","false")
Case "y","true","yes":o=IIf(IsTrue(o),"true","false")
Case "n","false","no","not":o=IIf(o<>"" And o<>"-" And(Not IsTrue(o)),"true","false")
Case "iz","isz","iszero":o=IIf(o<>"" And VOZ(o)=0,"true","false")
Case "uppercase","upper","u","ucase":o=UCase(o)
Case "lowercase","lower","l","lcase":o=LCase(o)
Case "propercase","proper","pcase","p":o=PCase(o)
Case "en","encode":o=Encode(o)
Case "en","decode":o=Decode(o)
Case "se","sqlen","sqlencode":o=UrlEncode(Replace(o,"'","''"))
Case "ue","urlen","urlencode":o=UrlEncode(o)
Case "ud","urlde","urldecode":o=URLDecode(o)
Case "es","escape":o=Escape(o)
Case "unes","unescape":o=Unescape(o)
Case "tags","notags":o=Replace(o,"[","[?")
Case "os","ostag":o=OSTag(o)
Case "ua","uatag":o=UATag(o)
Case "ep","eptag":o=EPTag(o)
Case "epdomain":o=EPDomain(o)
Case "url","uri":o=FileName(o)
Case "domain":
If o="$email$" Then
o=shop.CFG("email")
If o="" Then o=site.CFG("email")
If o="" Then o="info@"&MyDomain
Else
o=parser.Parse(o)
End If
o=rxSelect(o,"[\w\-\.]*$")
Case "domainid":
If o="$email$" Then
o=shop.CFG("email")
If o="" Then o=site.CFG("email")
If o="" Then o="info@"&MyDomain
Else
o=parser.Parse(o)
End If
o=rxSelect(o,"[\w\-\.]*$")
o=UCase(Left(o,6))
Case "txt","h2t","nohtml","striphtml","stripmarkup","textonly","plain","plaintext":o=StripHtml(o)
Case "htm","html","t2h","t2htm","t2html","text2html":o=Text2Html(o)
Case "text2lf","text2br","sc2br","sc2lf":o=rxReplace(o,"\s*\;\s*",vbCRLf)
Case "id","text","chars","char":o=rxReplace(o,"\W","")
Case "digits","numbers":o=rxReplace(o,"\D","")
Case "nosymbols":o=rxReplace(o,"[^\D\W]","")
Case "nonumbers":o=rxReplace(o,"[^\S\W]","")
Case "code":o=rxReplace(o,"[^\w\d\-\_]","")
Case "length","len","size":o=Len(o)
Case "empty","isempty","notset":o=LCase(CBool(o=""))
Case "set","isset","notempty":o=LCase(CBool(o<>""))
Case "zero","iszero","isz":o=LCase(CBool(VOZ(o)=0))
Case "nzero","notzero","nz":o=LCase(CBool(VOZ(o)<>0))
Case "abs","absolute":o=Abs(VOZ(o))
Case "is1","i1":o=LCase(CBool(VOZ(o)=1))
Case "gt1":o=LCase(CBool(VOZ(o)>1))
Case "lt1":o=LCase(CBool(VOZ(o)<1))
Case "ltz":o=LCase(CBool(VOZ(o)<0))
Case "gtz":o=LCase(CBool(VOZ(o)>0))
Case "gtez":o=LCase(CBool(VOZ(o)>=0))
Case "js","script":
o=EncodeJS(o)
Case "vb":
o=rxReplace(rxReplace(rxReplace(Encode(Replace(o,"""","""""")),"\r\n","""&vbCRLf&"""),"\r","""&vbCR&"""),"\n","""&vbLf&""")
Case "path":o=Replace(o,".","/")
Case "safepath":o=SafePath(o)
Case "archivepath","archive":o=ArchivePath(o)
Case "fullurl","permalink":o=FullyQualifiedUrl(o)
Case "ifie":o="<!--[?if IE]>"&o&"<![?endif]-->"
Case "cms","fwx","code":
o=Encode(rxReplace(o,"([\[\]\|\%\$\!\?\#])","(-(NoParse)-)$1(-(NoParse)-)"))
Case "quote","doublequote":o=""""&o&""""
Case "singlequote":o="'"&o&"'"
Case "brackets":o="("&o&")"
Case Else:
If rxTest(f,"^\w+$")Then
If False Then
ElseIf EXE("My_FWX__Format_"&f)Then
o=parser.ToBeFormatted
parser.ToBeFormatted=""
ElseIf EXE("FWX__Format_"&f)Then
o=parser.ToBeFormatted
parser.ToBeFormatted=""
End If
End If
End Select
End If
Next
s=o
Format=s
End Function
Public ToBeTranslated
Function Translate(ByVal s,ByVal f)
Me.ToBeTranslated=s
Dim o
o=Me.ToBeTranslated
f=LCase(Trim(rxReplace(f,"\W+"," ")))
For Each f In Split(f," ")
o=Application.Contents(f&":"&s)
If o<>"" Then
Else
If rxTest(f,"^\w+$")Then
If False Then
ElseIf EXE("My_FWX__Translate_"&f)Then
o=parser.ToBeTranslated
parser.ToBeTranslated=""
ElseIf EXE("FWX__Translate_"&f)Then
o=parser.ToBeTranslated
parser.ToBeTranslated=""
End If
End If
End If
Next
If o<>"" Then
s=Trim(o)
Else
s=o
End If
Translate=s
End Function
Public MergeText
Public RS
Public TagCache
Public iTag
Public aTag(10)
Private Sub Class_Initialize()
Me.MergeText="[\w\_\.\-\@\,\;\~\'\:\#]+"
Set Me.TagCache=Dictionary()
End Sub
Private Sub Class_Terminate()
Set Me.TagCache=Nothing
aTag=""
End Sub
End Class
Class fwxParserTag
Public Sub Class_Initialize()
Call Reset()
End Sub
Public Match
Public Data
Public ClassName
Public PropertyCount
Public Function TagID()
If Me.Data("tagid")="" Then Me.Data("tagid")=Me.TData("tagid,tagID,TagID")
If Me.Data("tagid")="" Then Me.Data("tagid")=MD5(Me.Match)
TagID=Me.Data("tagid")
End Function
Public Function TagKey()
If Me.Data("tagkey")="" Then Me.Data("tagkey")=Me.TData("tagkey,tagKey,TagKey,tagKEY,TagKEY,tag.key")
If Me.Data("tagkey")="" Then
Dim s:s=""
s=s&"TAG_"&Me.ClassName&""
If IsObject(Me.Source)Then s=s&"@"&Me.Source.PlainName&""
s=s&"#"&Me.TagID&""
Me.Data("tagkey")=s
End If
TagKey=Me.Data("tagkey")
End Function
Public Default Property Get D(ByVal f)
D=Data(f)
End Property
Public Property Let D(ByVal f,ByVal v)
Data(f)=v
End Property
Public Sub Reset()
With Me
.Match="":.ClassName=""
Set .Source=Nothing
Set .RS=Nothing
Set .Data=Nothing
Set .Data=Dictionary()
End With
End Sub
Public Function Start(ByVal sT)
With Me
.Reset()
.Match=sT
sT=rxReplace(sT,"^\[[\#\*]?","")
sT=rxReplace(sT,"\]$","")
.ClassName=sT
If rxTest(sT,"^\w+\s+")Then
sT=rxReplace(sT,"^(\w+)\s+","$1 ")
.ClassName=Mid(sT,1,InStr(1,sT," ")-1)
sConfig=Mid(sT,InStr(1,sT," ")+1,Len(sT))
.PropertyCount=UpdateCollection(Me.Data,sConfig)
If Not rxTest(.Data("args"),"(\w+)\=[\'\""]")Then
.Data("rawInput")=.Data("args")
End If
End If
End With
End Function
Public lclRS
Public Property Set RS(ByVal oRS)
Set lclRS=oRS
End Property
Public Property Get RS()
If Not IsObject(lclRS)Then
Set lclRS=Nothing
End If
If(lclRS Is Nothing)Then
If Not(Me.Source Is Nothing)Then
Set lclRS=Me.Source.RS()
End If
End If
If Err.Description<>"" Or Err.Number<>0 Then
DBG.FatalError "addins.tag.RS error: '"&Err.Description&"'"
End If
If(lclRS Is Nothing)Then
Else
Set RS=lclRS
End If
End Property
Public entCached
Public Property Set Source(ByVal ent)
Set entCached=ent
If Not(ent Is Nothing)Then
Me.Data("@K")=entCached.Key
If Me.Data("recurs")="" Then Me.Data("recurs")=entCached.DoIRecur()
End If
End Property
Public Property Get Source()
Dim entOutput
If IsObject(entCached)Then
If entCached Is Nothing Then
entCached=""
Else
Set entOutput=entCached
End If
End If
If Not IsObject(entOutput)Then
Dim strEntity
strEntity=Me.TData("datasrc,entity,ent,e,DataSrc,E,ENT,DATASRC")
If StartsWith(strEntity,"/")Then strEntity=""
If StartsWith(strEntity,"http:")Then strEntity=""
If StartsWith(strEntity,"https:")Then strEntity=""
If strEntity<>"" Then
If strEntity="module" Then
If j("c")="" Then
DBG.FatalError "ERR 65272: Module data source is undefined"
Else
Set entOutput=OBJ(j("c"))
If Err.Description<>"" Or Err.Number<>0 Then
DBG.FatalError "ERR 838373s5"
Set entOutput=Nothing
End If
End If
Else
Err.Clear()
Set entOutput=OBJ2(strEntity)
If Err.Description<>"" Or Err.Number<>0 Then
If IsBack Then
DBG.FatalError "Invalid Entity String '"&strEntity&"' ("&Err.Description&")"
End If
entOutput=""
End If
End If
If Err.Description<>"" Or Err.Number<>0 Then entOutput=""
End If
If IsObject(entOutput)Then
Set entCached=entOutput
Else
If IsFront Then
Set entOutput=site.Pages
Else
If Me.TData("xml,picasa,json")<>"" Or Me.ClassName="json" Then
Else
DBG.FatalError "Entity not specified in tag "&Me.ClassName&": "&Me.Match
End If
End If
End If
If IsObject(entOutput)Then
Me.Data("@K")=entOutput.Key
If Me.Data("recurs")="" Then Me.Data("recurs")=entOutput.DoIRecur()
End If
End If
If IsObject(entOutput)Then
Set Source=entOutput
Else
Set Source=Nothing
End If
End Property
Function Process_(ByVal s)
Process_=Eval(s)
End Function
Function Process()
If Left(Me.ClassName,6)="CDATA[" Then
Process=False
Exit Function
End If
If Left(Me.ClassName,5)="SKIP_" Then
Process=True
Me.Data("_output")=""
Exit Function
End If
If IsTrue(Me.BData("debugThenDie,debugAndDie,andDie,thenDie,thenDIE,die,thendie,dbgdie,DBGDIE"))Or IsTrue(Me.BData("debugThenDie,debugAndDie,dbgdie,debug"))Then
DBG.Show=True
DBG "###################"
DBG "################### ON-DEMAND TAG DEBUGGING..."
DBG "###################"
DBG "TAG CODE > "&Encode(Me.Match)&""
DBG "###################"
DBG "###################"
DBG "DATA RECEIVED > "&PrintCol(Me.Data)&""
DBG "###################"
DBG "###################"
End If
Me.Data("IF")=Me.TData("IF,If,if")
Me.Data("orIF")=Me.TData("orIF,orIf,orif")
If Me.Data("IF")<>"" Then
If GTZ(Me.Data("IF"))Or IsTrue(Me.Data("IF"))Or GTZ(Me.Data("orIF"))Or IsTrue(Me.Data("orIF"))Then
Else
Me.Data("_output")=Me.TData("IFn,IFfailed,IFnotalt,IFalt")
Exit Function
End If
End If
Me.Data("IFnot")=Me.TData("IFnot,Ifnot,ifnot,notIF,NotIf,notIf")
If Me.Data("IFnot")<>"" Then
If GTZ(Me.Data("IFnot"))Or IsTrue(Me.Data("IFnot"))Then
Me.Data("_output")=Me.TData("IFn,IFfailed,IFnotalt,IFalt")
Exit Function
Else
End If
End If
If Me.Data("NoCache")="" Then
If IsObject(MyCookies)Then
Me.Data("NoCache")=IsTrue(MyCookies.Global("NoCache.Flag")="y")
Else
Me.Data("NoCache")=IsTrue(Request.Cookies("NoCache.Flag")="y")
End If
End If
If IsFront And IsTrue(Me.Data("NoCache"))Then
ElseIf Me.Data("fresh")<>"" Then
Else
Me.Data("cache")=Me.TData("cache,Cache,CACHE")
If Me.Data("cache")<>"" Then
DBG "TAG-CACHE: '"&Me.Data("cache")&"'"
If IsNumeric(Me.Data("cache"))Then
Me.Data("cache")=VOZ(Me.Data("cache"))
If Me.Data("cache")="0" Then Me.Data("cache")=""
End If
If GTZ(Me.Data("cache"))Then
Me.Data("cache.key")=Me.TagID
Else
Me.Data("cache.key")=Me.Data("cache")
End If
If Me.Data("cache.key")<>"" Then
Me.Data("cache.uid")="#TAG-"&Me.Data("cache.key")
Me.Data("cache.fso")=MyDomainCache&"/TAGS/"&Me.ClassName&"/"&Me.Data("cache.key")&".txt"
DBG "TAG-CACHE: key="&Me.Data("cache.key")
DBG "TAG-CACHE: uid="&Me.Data("cache.uid")
DBG "TAG-CACHE: fso="&Me.Data("cache.fso")
If Me.Data("cache.content")="" Then
Me.Data("cache.content")=j(Me.Data("cache.uid"))
End If
If Me.Data("cache.content")="" And IsTrue(Me.TData("cache-ram,cache-in-ram"))Then
Me.Data("cache.content")=FWX_Cache_FromRAM(Me.Data("cache.uid"))
End If
If Me.Data("cache.content")="" Then
Me.Data("cache.content")=FWX_Cache_FromFILE(Me.Data("cache.fso"))
End If
Me.Data("_output")=Me.Data("cache.content")
DBG "TAG-CACHE: READ - content="&Len(Me.Data("cache.content"))&" bytes"
If Me.Data("_output")<>"" Then
DBG "TAG-CACHE: READ - using cache..."
Me.Data("cache.use")="y"
Else
DBG "TAG-CACHE: READ - cache was expired or not found..."
Me.Data("cache.content")=""
Me.Data("cache.put")="y"
DBG "TAG-CACHE: READ - updatah="&Me.Data("cache.put")
End If
End If
DBG "TAG-CACHE: _output="&Len(Me.Data("_output"))&"b"
End If
End If
If IsTrue(Me.Data("cache.use"))Then
Else
With Me
.Data("_tag")=.ClassName&""
.Data("_tag_name")="tag_"&.ClassName&""
.Data("_output")=""
If IsTest Then
End If
Me.Data("_tag_result")=Process_(.Data("_tag_name"))
If .Data("_output")<>"" Then
If Left(.Data("_output"),5)="file:" Then
.Data("_output")=ReadFile(Mid(.Data("_output"),5)&"")
End If
If Left(.Data("_output"),4)="url:" Then
.Data("_output")=GetResponse(Mid(.Data("_output"),4)&"")
End If
.Data("thenFormat")=.TData("thenFormat,then-format,postFormat,post-format")
If .Data("thenFormat")<>"" Then
.Data("_output")=parser.Format(.Data("_output"),.Data("thenFormat"))
End If
.Data("thenTranslate")=.TData("thenTranslate,then-translate,translate,language,locale")
If .Data("thenTranslate")<>"" Then
.Data("_output")=parser.Translate(.Data("_output"),.Data("thenTranslate"))
End If
.Data("outTemplate")=.TData("outTemplate,out-template,outputTemplate,output-template,outputAs")
If .Data("outTemplate")<>"" Then
.Data("_output")=MergeTemplate(.Data("outTemplate"),.Data)
End If
If IsSet(.Data("_output"))Then
.Data("_output")=.ParsedData("before")&.Data("_output")&.ParsedData("after")&""
If IsSet(.Data("wrap"))Then
wrap=Split(.Data("wrap"),"*")
If UBound(wrap)=1 Then
.Data("_output")=wrap(0)&.Data("_output")&wrap(1)&""
End If
End If
Dim ms,mo
If IsSet(.Data("merge"))Then
ms=Trim(.Data("merge"))
mo=Trim(.Data("_output"))
ms=Replace(ms,"|*|",mo)
ms=Replace(ms,"|@|",mo)
ms=Replace(ms,"|!|",mo)
.Data("_output")=TOT(ms,mo)
End If
.Data("_output")=Replace(.Data("_output"),"+equals+","=")
.Data("_output")=Replace(.Data("_output"),"+quotes+","""")
End If
End If
If .Data("_output")="" Then
If(Not rxTest(Me.ClassName,"^(DISABLED|SKIP|NOT)"))Then
.Data("_output")=.ParsedData("alt")
End If
End If
If .Data("_output")<>"" Then
If GTZ(.Data("LEFT"))Then
.Data("_output")=Left(.Data("_output"),VOZ(.Data("LEFT")))
End If
If GTZ(.Data("RIGHT"))Then
.Data("_output")=Right(.Data("_output"),VOZ(.Data("SLICE")))
End If
If GTZ(.Data("SLICE"))Then
.Data("_output")=Slice(.Data("_output"),VOZ(.Data("SLICE")))
End If
End If
End With
End If
If Me.Data("cache.put")<>"" Then
Me.Data("cache.content")=FWX_Cache_Payload_Make(Me.Data("_output"),TOT(Me.Data("cache"),TOT(Me.Data("cache.uid"),"tag-cache")))
DBG "TAG-CACHE: SAVE - uid="&Me.Data("cache.uid")&""
DBG "TAG-CACHE: SAVE - fso="&Me.Data("cache.fso")&""
DBG "TAG-CACHE: SAVE - content="&Len(Me.Data("cache.content"))&" bytes"
j(Me.Data("cache.uid"))=Me.Data("cache.content")
If Me.Data("cache.fso")<>"" Then
WriteToFile Me.Data("cache.fso"),Me.Data("cache.content")
End If
If IsTrue(Me.TData("cache-ram,cache-in-ram"))Then
Application.Contents(Me.Data("cache.uid"))=Me.Data("cache.content")
End If
DBG "TAG-CACHE: SAVED."
End If
If IsTrue(Me.BData("debugThenDie,debugAndDie,dbgdie,debug"))Then
DBG "###################"
DBG "################### OUTPUT = "&Encode(Me.data("_output"))
DBG "################### ON-DEMAND TAG DEBUGGING END"
DBG "###################"
DBG ""
DBG.Show=False
End If
If IsTrue(Me.BData("debugThenDie,debugAndDie,andDie,thenDie,thenDIE,die,thendie,dbgdie,DBGDIE"))Then
DBG.Show=True
DBG "###################"
DBG "################### DYING BY DEMAND AFTER TAG EXECUTION"
DBG "################### OUTPUT WAS: "&Encode(Me.Data("_output"))
DBG "################### TAG RUN-TIME: <br/>"&PrintCol(Me.Data)
DBG "###################"
DIE "END"
End If
If Me.TData("LOG,LOGIF")<>"" Then
If Me.Data("LOGIF")="" Then
Me.Data("LOGIF")="y"
End If
If Len(Me.Data("LOG"))>1 Then
Me.Data("name")=Me.TData("name,LOG")
End If
Dim li:li=Split(Me.Data("LOGIF"),"=")
If UBound(li)=1 Then
Me.Data("LOG")=rxTest(li(0),"("&Replace(li(1),",","|")&")")
Else
If Me.Data("LOGIF")<>"" Then Me.Data("LOGIF")=IIf(IsTrue(Me.Data("LOGIF")),"y","n")
End If
If IsTrue(Me.Data("LOGIF"))Then
Call tag_log()
End If
End If
End Function
Public Function tag__DATA_Initialize()
If Me.Data("done-tag__DATA_Initialize")<>"" Then Exit Function
Me.Data("done-tag__DATA_Initialize")="y"
If Me.Data("recurs")="" Then Me.Data("recurs")=Me.Source.DoIRecur()
Me.Data("recurs")=IsTrue(Me.Data("recurs"))'IIf(Me.Data("recurs"),"y","n")
If IsTrue(Me.Data("all"))Then
Me.Data("max")=-1
Me.Data("@size")=-1
Me.Data("@page")=1
End If
Me.Data("max")=Me.NData("max,limit,maximum,top")
Me.Data("@size")=Me.NData("@size,recpp,PageSize,size")
Me.Data("@page")=VOV(Query.FQ("@page"),Me.Data("@page"))
If Not Me.Data("@page")>0 Then Me.Data("@page")=1
Me.Data("nodiv")=Me.TData("nodiv,noDiv,NoDiv,NODIV,raw")
Me.Data("child")=Me.TData("child,children")
Me.Data("directory")=CBool(Me.BData("directory,dir")Or IsSet(Me.Data("child")))
If Me.Data("directory")Then
Me.Data("childParent")=Me.Data("@K")
If Not GTZ(Me.Data("childMax"))Then Me.Data("childMax")=3
End If
Dim sF:sF=""
sF=Me.ParsedData("filter")
sF=Enclose(sF,"(",")")
If False Then
ElseIf Me.Data("active")="-" Then
ElseIf InStr(1,sF,"!search")>0 Then
ElseIf InStr(1,sF,"%SEARCH")>0 Then
ElseIf InStr(1,sF,"%FILTER")>0 Then
ElseIf InStr(1,sF,"[active]")>0 Then
Else
sF=MergeCriteria(sF,"AND","/*(tag.data|entity).active*/ISNULL(t.[active],0)="&IIf(TOT(Me.Data("active"),"y"),"1","0"))
End If
sF=TOT(Me.ParsedData("where"),sF)
If InStr(1,sF,"!search")>0 Then
sF=Replace(sF,"!search","%SEARCH")
Else
sF=EncloseCriteria(MergeCriteria(sF,"AND",TOT(Source.Filter,"1=1")))
End If
If(Me.Data("module")<>"")Then
sF=MergeCriteria(sF,"AND","t.[module]='"&Me.Data("module")&"'")
End If
Me.Data("list")=TOT(TOT(Me.Data("list"),Me.Data("publish")),Me.Data("published"))
If IsFront And Me.Data("list")="" And Me.Data("unlisted")="" Then
If rxTest(Source.Key,"(page|product|article|download|link|chapter|faq|image|feature|spec)")Then Me.Data("list")="yes"
End If
If(Me.Data("list")<>"")And(Me.Data("list")<>"-")Then
sF=MergeCriteria(sF,"AND","(t.[list] = "&IIf(IsTrue(Me.Data("list")),1,0)&")")
End If
If IsFront And Me.Data("menu")="" And Me.Data("_tag")="menu" Then
If rxTest(Source.Key,"(page|product|article)")Then Me.Data("menu")="yes"
End If
If(Me.Data("menu")<>"")And(Me.Data("menu")<>"-")Then
sF=MergeCriteria(sF,"AND","(t.[menu] = "&IIf(IsTrue(Me.Data("menu")),1,0)&")")
End If
Me.Data("special")=TOT(TOT(Me.Data("special"),Me.Data("highlight")),Me.Data("featured"))
If(Me.Data("special")<>"")And(Me.Data("special")<>"-")Then
sF=MergeCriteria(sF,"AND","(t.[special] = "&IIf(IsTrue(Me.Data("special")),1,0)&")")
End If
Me.Data("frontpage")=Me.TData("frontpage,front")
If(Me.Data("frontpage")<>"")And(Me.Data("frontpage")<>"-")Then
sF=MergeCriteria(sF,"AND","(t.[frontpage] = "&IIf(IsTrue(Me.Data("frontpage")),1,0)&")")
End If
If(Me.Data("search")<>"")And(Me.Data("search")<>"-")Then
sF=MergeCriteria(sF,"AND","(t.[search] = "&IIf(IsTrue(Me.Data("search")),1,0)&")")
End If
If(Me.Data("unlisted")<>"")Then
sF=MergeCriteria(sF,"AND","t.[list] = 0")
End If
If(Me.Data("clearance")<>"")Then
sF=MergeCriteria(sF,"AND","t.[clearance] = 1")
End If
Me.Data("filter")=sF
Me.Data("sql")=Me.TData("sql,Sql,SQL")
Me.Data("nlist")=Me.TData("nlist,olist,ids,idlist,ilist")
If Me.Data("nlist")<>"" Then
Me.Data("nlist")=NList(Me.Data("nlist"))
Me.Data("filter")=MergeCriteria("t.%KEY IN ("&Me.Data("nlist")&")","AND",Me.Data("filter"))
Me.Data("fw")="y"
End If
Me.Data("sort")=Me.TData("sort,order")
If Me.Data("sort")="" Then
If Me.Data("nlist")<>"" Then
Me.Data("sort")=rxReplace(Me.Data("sort"),"^(ORDER BY)","$1 dbo.fncSortInto(t.%KEY,'"&Me.Data("nlist")&"') ASC,")
End If
If Me.Data("sort")="" Then
Me.Data("sort")=Me.Source.Sort()
End If
End If
If Me.Data("sort")<>"" Then
Me.Data("sort")=rxReplace(Me.Data("sort"),"^\s*ORDER\s*BY\s*","")
If Left(Me.Data("sort"),1)="!" Then
Me.Data("sort")=Replace(Me.Data("sort"),"!","")&""
End If
If Me.Data("sort")<>"" Then
If Me.Data("sort")="none" Then
Me.Data("sort")=""
Else
Me.Data("sort")="ORDER BY "&Me.Data("sort")&""
End If
End If
End If
Dim sRF:sRF=""
If IsObject(Me.Source)Then
If False And Me.Data("root-filter")<>"" Then
Else
If Me.Data("fk")<>"" Or Me.Source.Parent<>"" Then
Me.Data("root")=TOT(Me.TData("root,base"),Me.Source.ParentID)
Me.Data("fw")=IIf(Me.BData("wild,fw,Wild,WildMatch"),"y","n")
If Me.Data("root")<>"" Or Me.BData("recurs,fw")Then
Me.Data("root")=VOZ(Me.Data("root"))
If Me.Data("ff")="invoice" And Me.Data("ft")<>"" Then
sRF=MergeCriteria(sRF,"AND","/*ROOT-INV*/"&Me.Data("ft")&".["&Me.Data("fm")&"]="&Me.Data("root")&"")
Else
If Me.Data("root")>0 Then
Me.Data("ft")=TOT(Me.TData("ft,folder-table,browse-table,root-table"),"t")
Me.Data("fk")=TOT(Me.TData("fk,folder-key,browse-key,root-key"),Me.Source.Parent)
Me.Data("ff")=TOT(Me.TData("ff,folder-field,browse-field,root-field"),Me.Data("fk"))
Me.Data("fm")=TOT(Me.TData("fm,folder-map,browse-map,root-map"),IIf(Me.Data("recurs"),"location",Me.Data("fk")&"Map"))
If Me.Data("ft")="f" And LCase(Me.Data("f")="site.pages")And Me.Data("fm")="pageMap" Then
Me.Data("fm")="location"
End If
If Right(Me.Data("ft"),1)="." Then
Me.Data("ft")=Replace(Me.Data("ft"),".","")
End if
If IsTrue(Me.Data("fw"))And IsTrue(Me.Source.ParentRecurs)And(Not IsTrue(Me.Data("no-root-map")))And(Not IsTrue(Me.Data("cant-be-wild")))Then
If IsSet(Me.Data("fw"))And Me.Data("root")=0 Then
sRF="(301=301)/*ROOT-N*/"
ElseIf Me.Data("root")=0 Then
sRF="(103=103)"'/*ROOT-E=any*/)"
Else
If IsTrue(Me.Data("cant-be-wild"))Or IsTrue(Me.Data("no-root-map"))Then
sRF=MergeCriteria(sRF,"AND","/*ROOT-ARGH*/ISNULL("&Me.Data("ft")&".["&Me.Data("ff")&"],0)="&VOZ(Me.Data("root"))&"")
Else
sRF=MergeCriteria(sRF,"AND","/*ROOT-A*/ISNULL("&Me.Data("ft")&".["&Me.Data("fm")&"],'{0}') LIKE '%{"&Me.Data("root")&"}%'")
End If
If Me.Data("recurs")Then
sRF=MergeCriteria(sRF,"AND","/*ROOT-B*/t.["&Me.Data("@K")&"]<>"&Me.Data("root")&"")
End If
End If
Else
sRF=MergeCriteria(sRF,"AND","/*ROOT-C*/"&Me.Data("ft")&".["&Me.Data("ff")&"] = "&Me.Data("root")&"")
End If
End If
End If
Else
sRF="(101=101)"'/*ROOT-EA=no root*/)"
End If
Else
sRF="(102=102)"'/*ROOT-EB=no parent field*/)"
End If
End If
Dim tagw,tags
Me.Data("tag")=TOT(Me.TData("tag,tags"),"")
If Me.Data("tag")<>"" Then
tagw=TOT(Me.TData("tag-filter,tag-ftr"),"OR")
tags=TOT(Me.TData("tag-list,tag-field"),"tags")
sRF=MergeCriteria(sRF,tagw,"dbo.InList('"&SQLSafeString(Me.Data("tag"))&"',"&tags&")=1")
End If
Me.Data("parents")=TOT(Me.TData("parents,roots"),"")
If Me.Data("parents")<>"" Then
tagw=TOT(Me.TData("parents-filter,parents-ftr"),"OR")
tags=TOT(Me.TData("parents-list,parents-field"),"parents")
sRF=MergeCriteria(sRF,tagw,"dbo.InList('"&SQLSafeString(Me.Data("parents"))&"',"&tags&")=1")
End If
Else
sRF=""'(1=1/*ROOT-EC=no entity found*/)"
End If
Me.Data("root-filter")=TOT(Enclose(sRF,"(",")"),Me.Data("root-filter"))
End Function
Public Function tag__DATA_Prepare_Template()
If Me.Data("done-tag__DATA_Prepare_Template")<>"" Then Exit Function
Me.Data("done-tag__DATA_Prepare_Template")="y"
tag__DATA__FindTemplate ""
End Function
Public Function tag__DATA__FindTemplate(ByVal prfx)
Dim o:o=""
Me.Data(prfx&"templateFile")=Me.TData(prfx&"templateFile,"&prfx&"TemplateFile,"&prfx&"templatefile")
Me.Data(prfx&"getTemplate")=Me.TData(prfx&"getTemplate,"&prfx&"GetTemplate,"&prfx&"gettemplate")
Me.Data(prfx&"template")=Me.TData(prfx&"template,"&prfx&"Template,"&prfx&"tmpl,"&prfx&"t,"&prfx&"T,"&prfx&"tpl")
If o="" And Me.Data("cache-template")<>"" Then
If Left(Me.Data("cache-template"),1)="/" Then
Me.Data("cache-template-file")=Me.Data("cache-template")
Else
Me.Data("cache-template-file")="/$$TEMP/"&MyDomain&"/TAGS/listing-templates/"&Me.Data("cache-template")&".txt"
End If
o=j("#TAG-"&Me.Data("cache-template-file"))
If o="" Then
o=FWX_Cache_FromFILE(Me.Data("cache-template-file"))
If o<>"" Then
j("#TAG-"&Me.Data("cache-template-file"))=o
Me.Data("cache-template-write")="y"
End If
Else
End If
End If
If o="" Then
Set o=New QuickString
With o
If 1=0 Then
ElseIf IsSet(Me.Data(prfx&"template"))Then
If 1=0 Then
ElseIf LCase(Me.Data(prfx&"template"))="headlines" Then
.Add "<a class=""ListTitle"" href='[tot a=""|redirect|"" b=""|link|"" b=""|url|""]' style=""width:100%;"">"
.Add "<b class=""NoUnderline"">[24hr |date|]</b>[hsp 5][tot a=""|link_text|"" b=""|name|"" c=""|title|""]"
.Add "</a>"
ElseIf LCase(Me.Data(prfx&"template"))="news" Then
.Add "<a class=""ListTitle"" href='[tot a=""|redirect|"" b=""|link|"" b=""|url|""]' style=""width:100%;"">"
.Add "<span class=""NotU"">[dodate |date|]</span>[hsp 5]<b>[tot a=""|link_text|"" b=""|name|"" c=""|title|""]</b>"
.Add "</a>"
.Add "<div class=""ListDescription"">"
.Add "[slice a=""|intro|"" b=""|description|"" length=""160""]"
.Add "</div>"
ElseIf LCase(Me.Data(prfx&"template"))="images" Then
.Add "<h3><a href='[tot a=""|redirect|"" b=""|link|"" b=""|url|""]'>|name|</A></h3>"
.Add "<img src=""|image|"" width=120><br />"
.Add "<a href=""|link|"">more info >></A>"
ElseIf LCase(Me.Data(prfx&"template"))="browse" Then
Me.Data("nodiv")="y"
If LCase(Me.ClassName)="addressbar" Then
.Add "&nbsp;<a href=""[resubmit "&Me.Data("@K")&"=|"&Me.Data("@K")&"|]"">|name|</A>&nbsp;"
Else
.Add "<a href=""[resubmit "&Me.Data("@K")&"=|"&Me.Data("@K")&"|]"">|name|</A>"
End If
ElseIf LCase(Me.Data(prfx&"template"))="menu" Then
Me.Data("nodiv")="y"
o="<li><a href=""[tot a=""|redirect|"" b=""|link|"" c=""|url|""]"" title=""[tot a=""|tooltip|"" b=""|title|"" c=""|description|""]"">[tot a=""|link_text|"" b=""|name|"" c=""|title|""]</a></li>"
ElseIf LCase(Me.Data(prfx&"template"))="options" Then
Me.Data("nodiv")="y"
o="<option value=""[tot a=""|value|"" b=""|id|"" c=""|key|""]""[selected a=""|value|"" b=""|id|"" c=""|key|"" current=""|current|""]>[tot a=""|title|"" b=""|name|"" c=""|value|""]</option>"
ElseIf rxTest(LCase(Me.Data(prfx&"template")),"^(links?|breadcrumbs?|trail|addressbar)$")Then
Me.Data("nodiv")="y"
o="<a href='[tot a=""|redirect|"" b=""|link|"" b=""|url|""]' title=""[tot a='|tooltip|' b='|description|']"">[tot a=""|link_text|"" b=""|name|"" c=""|title|""]</a>"
Else
o=Me.Data(prfx&"template")
End If
ElseIf IsSet(Me.Data(prfx&"vbTemplate"))Then
o=EXE(Me.Data(prfx&"vbTemplate"))
ElseIf IsSet(Me.Data(prfx&"getTemplate"))Then
Me.Data(prfx&"getTemplate")=rxReplace(Me.Data(prfx&"getTemplate"),"\<[^\>]*\>","")
o=EXE(TOT(j("t"),IIf(IsFront,"FE__TMPL","BO__TMPL"))&"("""&Me.Data(prfx&"getTemplate")&""")")
ElseIf IsSet(Me.Data(prfx&"file"))Then
o=ReadFile(Me.Data(prfx&"file"))
ElseIf IsSet(Me.Data(prfx&"templateFile"))Then
o=TemplateFile(Me.Data(prfx&"templateFile"))
ElseIf prfx="" And Me.ClassName="menu" Then
o="[tot a=""|label|"" b=""|link_text|"" c=""|name|""]"
ElseIf prfx="" And rxTest(Me.ClassName,"addressbar|trail|breadcrumbs")Then
Me.Data("nodiv")="y"
o="<a href=""[tot a=""|redirect|"" b=""|link|"" b=""|url|""]"" title=""[tot a=""|tooltip|"" b=""|description|""]"">[tot a=""|link_text|"" b=""|name|"" c=""|title|""]</a>"
ElseIf prfx="" Then
o="[tot a=""|name|"" b=""|title|""]"
End If
End With
If o<>"" And Me.Data("cache-template-write")<>""<>"" Then
WriteToFile Me.Data("cache-template-file"),FWX_Cache_Payload_Make(o,"tag-data-list-item-cache")
End If
End If
Me.Data(prfx&"template")=o
End Function
Public Function tag__DATA_MergeWithObject(ByRef rsSource)
If IsEmpty(Me.Data("template"))Then
DBG "missing template, loading emergency template"
Me.Data("template")="<a href='[tot a=""|redirect|"" b=""|link|"" b=""|url|""]'>[tot a='|link_text|' b='|name|']</a>[vsp 1]|description|"
End If
output=MergeTemplate(Me.Data("template"),rsSource)
tag__DATA_MergeWithObject=output
End Function
Public Function tag__DATA_SQL()
If Me.Data("done-tag__DATA_SQL")<>"" Then Exit Function
Me.Data("done-tag__DATA_SQL")="y"
Me.Data("criteria")=MergeCriteria(Me.Data("filter"),"AND",Me.Data("root-filter"))
Me.Data("criteria")=Me.TData("where,criteria")
If Me.Data("fields")<>"" Then
Me.Data("fields")=Me.ParsedData("fields")
End If
If Me.Data("fields")="%FIELDS" Then
ElseIf Me.Data("fields")<>"" Then
Dim f,fs
For Each f In Split(Me.Data("fields"),",")
f=Trim(f)
If Instr(1,f,".")>0 Or Instr(1,f," ")>0 Then
ElseIf Instr(1,f,"@")>0 Or Instr(1,f,"%")>0 Then
ElseIf StartsWith(LCase(Trim(f)),"convert(")Then
ElseIf StartsWith(LCase(Trim(f)),"datepart(")Then
ElseIf rxTest(f,"[\r\n\t]")Then
ElseIf InStr(Trim(f),"t.")Then
Else
f="t."&f
End If
fs=PutAfter(fs,",")&f
Next
Me.Data("fields")=fs
If Me.Data("group")="" Then
Me.Data("fields")="t.%KEY,"&Me.Data("fields")
End If
If Me.Data("fields")<>"" Then
If InStr(1,Source.FullPlainName,"stock")Then
ElseIf Source.Key="product" Then
f=Me.Data("fields")
If Not rxTest(f,"fncINVY_Product_Stock")Then f=rxReplace(f,"(\w+\.)?\[?liveStock\]?","dbo.[fncINVY_Product_Stock](%QUOTESTOCKIST, $1[product], DATEADD(m,6/*pre-booking-months*/,'"&DoDate(DOD(Query.FQ("asof"),Now()))&"')) AS 'liveStock'")
If Not rxTest(f,"fncINVY_Product_HardStock")Then f=rxReplace(f,"(\w+\.)?\[?hardStock\]?","dbo.[fncINVY_Product_HardStock](%QUOTESTOCKIST, $1[product], '"&DoDate(DOD(Query.FQ("asof"),Now()))&"') AS 'hardStock'")
If Not rxTest(f,"fncINVY_Product_Trade_Price")Then f=rxReplace(f,"(\w+\.)?\[?tradePrice\]?","dbo.[fncINVY_Product_Trade_Price](%QUOTEMERCHANT, $1[product], $1[fixprice]) AS 'tradePrice'")
Me.Data("fields")=f
End If
If Source.Key="option" Then
Me.Data("fields")=Replace(Me.Data("fields"),"tradePrice","dbo.[fncINVY_Service_Option_Trade_Price](%QUOTEMERCHANT, [option], [fixprice]) AS 'tradePrice'")
End If
End If
Me.Data("fields")=Me.Data("fields")&Enclose(Me.TData("extraFields,moreFields,extrafields,morefields"),",","")
End If
Me.Data("--sql")="SELECT "&IIf(GTZ(Me.Data("max")),"TOP "&Me.Data("max")&" ","")&TOT("%FIELDS",Me.Data("fields"))&" FROM "&TOT(Me.TData("sql-table,sql-from,-table,-from"),"%TABLE %NOLOCK")&" "&TOT(Me.TData("sql-join,-join"),"%JOIN")&" WHERE "&Me.TData("criteria")&" "&TOT(Me.TData("sql-group,-group"),"%GROUP")&" "&Me.Data("sort")&""
tag__DATA_SQL=Me.Data("--sql")
End Function
Public Function tag_address():tag_address=tag_breadcrumbs:End Function
Public Function tag_addressbar():tag_addressbar=tag_breadcrumbs:End Function
Public Function tag_breadcrum():tag_breadcrum=tag_breadcrumbs:End Function
Public Function tag_breadcrums():tag_breadcrums=tag_breadcrumbs:End Function
Public Function tag_breadcrumb():tag_breadcrumb=tag_breadcrumbs:End Function
Public Function tag_breadcrumbs():tag_breadcrumbs=True
Call tag__DATA()
Me.Data("root")=Me.NData("root,base,current,id")
Me.Data("current")=Me.NData("current,cur,root")
Me.Data("trail")=Me.TData("trail,location,loc,track,map")
If Me.Data("trail")="" Then
If Me.Data("root")>0 Then
Me.Data("trail")=Me.Source.FieldByID("location",Me.Data("root"))
End If
If Me.Data("trail")="" Then
Exit Function
End If
End If
Me.Data("trail")=TOE(Me.Data("trail"))
If Me.Data("trail")="" Then
Exit Function
End If
Me.Data("trail")=Replace(Me.Data("trail"),"}{",",")
Me.Data("trail")=Replace(Me.Data("trail"),"{","")
Me.Data("trail")=Replace(Me.Data("trail"),"}","")
If Me.Data("trail")="0" Then
Exit Function
End If
If Me.Data("template")="" Then
Me.Data("template")="<a href='[tot a=""|redirect|"" b=""|link|"" c=""|url|""]'>[tot a=""|link_text|"" b=""|name|""]</a>"
End If
If Not IsTrue(Me.Data("NoDivider"))Then
sDivider=Me.Data("divider")
If sDivider="" Then sDivider="&nbsp;&raquo;&nbsp;"
End If
If VOZ(Me.Data("spacing"))>0 Then sSpacing=hsp(Me.Data("spacing"))
Me.Data("@K")=Source.Key
Dim sF:sF=""
If rxTest(Me.Source.Key,"page|product|article|event")Then
sF=",[link_text],[description],[link],[redirect]"
End If
Dim rsAB,sAB_SQL
sAB_SQL="/*Breadcrumbs*/SELECT [name],[parent],%KEY"&sF&"FROM %TABLE %NOLOCK WHERE %KEY IN ("&Me.Data("trail")&") %SORT"
Set rsAB=Me.Source.Run(sAB_SQL)
If Not Exists(rsAB)Then
Exit Function
End If
Dim aCrums,sCrum,iCrum,output
Set output=New QuickString
output.Divider=sDivider
aCrums=Split(Me.Data("trail"),",")
For Each sCrum In aCrums
iCrum=VOZ(sCrum)
DBG "TRY breadcrumb: "&iCrum
If iCrum>0 Then
rsAB.Filter="("&Me.Data("@K")&"="&iCrum&")"
If Exists(rsAB)Then
If iCrum<>Me.Data("hide")Then
If iCrum=Me.Data("current")And Me.Data("hideCurrent")<>"" Then
DBG "CURRENT breadcrumb: "&iCrum
Else
DBG "ADD breadcrumb: "&iCrum
output.Add tag__DATA_MergeWithObject(rsAB)&""
End If
End If
End If
End If
Next
KillRS rsAB
Me.Data("_output")=output
End Function
Public Function tag_Dropdown__NODE(ByVal iRoot,ByVal bRecur,ByVal sIndent)
Dim sL,sV,sP,sS,sC,sD,sI,rD
Dim output
Set output=New QuickString
With output
.Divider=""
Set rD=Me.Data("_DATA").Clone()
If bRecur And Me.Data("recurs")<>"n" Then
Me.Data("|r|")=Me.Data("|r|")&"|"&iRoot&"|"
rD.Filter="parent = "&iRoot&""
End If
If Err.Description<>"" Or Err.Number<>0 Then
DBG.FatalError "FATAL ERROR in 'tags.dropdown': '"&Err.Description&"'<br/>Recurs:"&Me.Source.DoIRecur()&"<br/>"&Encode(Me.Match)
Exit Function
End If
If Not Exists(rD)Then
Exit Function
Else
Do While Exists(rD)
If Me.Data("option-template")<>"" Then
.Add MergeTemplate(Me.Data("option-template"),rD)
Else
sL=Me.TData("label,label-template")
If sL<>"" Then
sL=MergeTemplate(sL,rD)
End If
If sL="gap" Or sL="spacer" Or sL="space" Then
.Add "<option></option>"&vbCRLf
Else
sV=MergeTemplate(Me.Data("value"),rD)
sP=MergeTemplate(Me.Data("properties"),rD)
sS=MergeTemplate(Me.Data("style"),rD)
sC=MergeTemplate(Me.Data("class"),rD)
sD=MergeTemplate(Me.Data("data"),rD)
sI=Replace(sIndent,".",Me.Data("indent"))
.Add "<option value="""&sV&""""
If sP<>"" Then .Add " "&sP&""
If sS<>"" Then .Add " style="""&sS&""""
If sC<>"" Then .Add " class="""&sC&""""
If sD<>"" Then .Add " data="""&sD&""""
.Add ">"
If sI<>"" Then .Add sI&" "
.Add sL&""
.Add "</option>"&vbCRLf
If Err.Description<>"" Or Err.Number<>0 Then
.Add "<option value="""&Encode(Err.Description)&""" onclick=""alert(this.value)"">Error #"&Err.Number&"</option>"
Err.Clear()
End If
If bRecur Then
iNewRoot=rD(Me.Data("_k"))
If InStr(1,Me.Data("|r|"),"|"&iNewRoot&"|")>0 Then
.Add "<option>No DUP '"&iNewRoot&"'</option>"&vbCRLf
Else
sList=tag_Dropdown__NODE(iNewRoot,True,sIndent&".")
If sList<>"" Then
.Add sList
End If
End If
End If
End If
End If
rD.MoveNext
Loop
End If
KillRS rD
End With
Err.Clear()
tag_Dropdown__NODE=output
End Function
Public Function tag_Dropdown():tag_Dropdown=True
If Not IsObject(Me.Source)Then Exit Function
If IsTrue(Me.Data("recurs"))Then Me.Data("wild")=True
Me.Data("root")=Me.NData("root")
If Me.Data("fields")="" Then
Me.Data("fields")="name,sort"
If Me.Source.Recurs Then
Me.Data("fields")=Me.Data("fields")&",parent,location"
End If
End If
Me.Data("fields")=Me.Data("fields")&Enclose(Me.TData("extraFields,moreFields,extrafields,morefields"),",","")
If Me.Data("cache")="" Then
If Me.Data("kache")="e" Then Me.Data("kache")=rxReplace(LCase(Me.TData("datasrc,entity,ent,e")&Enclose(Me.JData("where,filter",","),"{,","}")),"[^\w\d\-\.\_\{\}]+","")
If Me.Data("kache")="" Then Me.Data("kache")=Me.TagKey()
Else
Me.Data("kache")=""
End If
If Me.Data("kache")<>"" Then
Me.Data("kache")=rxReplace(Me.Data("kache"),"[^\w\d\-\.\_\{\}]+","")
End If
If Me.Data("kache")<>"" Then
If output="" Then
output=Application.Contents("#C-TAG-"&Me.Data("kache"))
If output<>"" Then
j("_STAT-dropdown-cache-ram")=VOZ(j("_STAT-dropdown-cache-ram"))+1
DBG "tag_DROPDOWN> CACHE FOUND (M): "&Len(output)&" bytes"'<br/>"&Encode(output)
End If
End If
If output="" Then
Me.Data("fache")=Me.Data("kache")
Me.Data("fache")=MyDomainCache&"/TAGS/"&Me.ClassName&"/"&Me.Data("fache")&"/E~"&j("EDITION")&"_R~"&j("REVISION")&".txt"
Me.Data("fache")=Replace(Me.Data("fache"),"TAG#@","")
Me.Data("fache")=Replace(Me.Data("fache"),"@","/")
output=ReadFileWithoutValidation(Me.Data("fache"))
If output<>"" Then
j("_STAT-dropdown-cache-fso")=VOZ(j("_STAT-dropdown-cache-fso"))+1
DBG "tag_DROPDOWN> CACHE FOUND (F): "&Len(output)&" bytes"
End If
End If
End If
If output="" Then
j("_STAT-dropdown-fresh")=VOZ(j("_STAT-dropdown-fresh"))+1
End If
If output="" Then
Call tag__DATA()
Dim rL
If IsObject(Me.Data("rs"))Then
Set rL=Me.Data("rs")
Else
If Me.Data("rs")<>"" Then
If Me.Data("rs")="!" Then
Set rL=j("rs")
Else
Set rL=OBJ2(Me.Data("rs"))
End If
End If
End If
If(Not IsObject(rL))Then
Me.Source.Filters("_customSTAKEHOLDER")=""
If Me.Data("sql")<>"" Then
Set rL=Me.Source.Run("/*DROPDOWN-sql*/"&Me.Data("sql"))
ElseIf Me.Source.Group<>"" Then
Set rL=Me.Source.Run("/*DROPDOWN-group*/SELECT "&IIf(GTZ(Me.Data("max")),"TOP "&VOZ(Me.Data("max"))&" ","")&TOT(Me.Data("fields"),"%FIELDS")&" FROM %TABLE %NOLOCK %JOIN WHERE "&Me.Data("filter")&" %GROUP "&Me.Data("sort")&"")
Else
Me.Data("sql")=Me.Data("--sql")
Me.Data("sql")=Replace(Me.Data("sql"),"%FIELDS",TOT(Me.Data("fields"),"%FIELDS"))
Set rL=Me.Source.Run(Me.Data("sql"))
End If
End If
If Exists(rL)Then
Set Me.Data("_DATA")=rL
If IsEmpty(Me.Data("indent"))Then Me.Data("indent")=Me.Data("IndentWith")
If IsEmpty(Me.Data("indent"))Then Me.Data("indent")="|__ "
If IsEmpty(Me.Data("indentStart"))Then Me.Data("indentStart")=Me.Data("StartIndent")
If IsEmpty(Me.Data("indentStart"))Then Me.Data("indentStart")=Me.Data("start")
If Me.Data("noIndent")<>"" Then Me.Data("indent")=""
If IsEmpty(Me.Data("label"))Then
Me.Data("label")="|name|"
End If
Me.Data("maxlength")=Me.NData("maxlength,max-length,maxLength")
If Not GTZ(Me.Data("maxlength"))Then
Me.Data("maxlength")=100
Else
Me.Data("maxlength")=VOZ(Me.Data("maxlength"))
End If
Me.Data("value")=Me.TData("value,val,option-value,optionValue,value-template")
If Me.Data("value")="" Then Me.Data("value")=Me.TData("valueField,value-field")
If Me.Data("value")="" Then Me.Data("value")=Me.Source.Key
If rxTest(Me.Data("value"),"^\w+$")Then
Me.Data("value")="|"&Me.Data("value")&"|"
End If
Me.tag__DATA__FindTemplate "option-"
Me.Data("_k")=Me.Source.Key
output=tag_Dropdown__NODE(Me.Data("root"),Me.Data("recurs"),Me.Data("indentStart"))
End If
output=TOT(output,"!EMPTY!")
If Me.Data("kache")<>"" Then
Application.Contents("#C-TAG-"&Me.Data("kache"))=output
If Me.Data("fache")<>"" Then WriteToFile Me.Data("fache"),output
End If
End If
If output="!EMPTY!" Then output=""
If output="" Then
output=""'<!--//EMPTY LIST//-->"
Else
If IsEmpty(Me.Data("notselected"))Then
Call tag_tot()
Me.Data("current")=Me.TData("current,selected,cur,def,string,default")
For Each val In Split(Me.Data("current"),",")
val=Trim(val)&""
output=Replace(output,"value="""&val&"""","value="""&val&""" selected=""selected"" ")
Next
End If
End If
Set Me.Data("_DATA")=Nothing
Me.Data("_output")=output
End Function
Public Function tag_list__PRINT(ByRef rL)
tag_list__PRINT=""
If Not Exists(rL)Then
Exit Function
End If
If Me.Data("template")="" Then
Me.Data("template")=Me.Data("getTemplate,GetTemplate,gettemplate")
Me.Data("template")=GetTemplate(Me.Data("template"))
End If
If Me.Data("template")="" Then
DBG.FatalError "Missing template for tag: "&Me.Match
End If
With Me
Dim be:be=.TData("before-each,beforeEach,beforeeach")
Dim ae:ae=.TData("after-each,afterEach,aftereach")
Dim we:we=.TData("wrap-each,wrapEach,wrapeach")
If we<>"" Then
If InStr(1,we,"*")>0 Then
Me.Data("template")=Replace(we,"*",Me.Data("template"))
Else
we=""
End If
End If
Me.Data("template")=Enclose(Me.Data("template"),be,ae)
End With
Me.Data("@last")=rL.RecordCount
Me.Data("startFrom")=VOV(Me.Data("startFrom"),1)
If Me.Data("max")<1 Then Me.Data("max")=rL.RecordCount
If VOZ(Me.Data("@size"))<0 And Me.Data("max")>0 Then Me.Data("@size")=VOZ(Me.Data("max"))
If VOZ(Me.Data("@size"))=0 Then Me.Data("@size")=FWX_DEFAULT_RECPP
If Me.Data("@page")<1 Then Me.Data("@page")=1
If Not GTZ(Me.Data("rows"))Then Me.Data("rows")=Me.Data("@size")
Me.Data("cells")=Me.NData("cells,cols,columns")
If Not GTZ(Me.Data("cells"))Then Me.Data("cells")=1
Me.Data("@top")=VOZ(Me.NData("@top,@size,max"))
Me.Data("directory")=TOT(Me.Data("directory"),Me.Data("children"))
Me.Data("count")=CBool(IsSet(Me.Data("count"))Or IsSet(Me.Data("directory")))And IsEmpty(Me.Data("nocount"))
If IsTrue(Me.Data("directory"))Then
Me.Data("class")=Me.Data("class")&" Directory"
Else
Me.Data("class")=Me.Data("class")&" List"
End If
If Me.Data("cells")>0 Then
Me.Data("class")=Me.Data("class")&" Cells Cells"&Me.Data("cells")&""
End If
Me.Data("class")=Me.Data("class")&" Clear"
Dim dL:Set dL=Dictionary()
UpdateCollection dL,Me.Data
dL("#k")=Me.Data("@K")
dL("#e")=Me.Source.PlainName'TOT(Me.Data("datasrc"),Me.Source.PlainName)
dL("#f")=Me.Data("filter")
dL("#sql")=Me.Data("sql")
Dim sListQS:Set sListQS=New QuickString
sListQS.Divider=Me.Data("divider")
Dim rsChildren,iCount
iCount=1:rL.MoveFirst
rL.PageSize=Me.Data("@size")
If Me.Data("@page")<=rL.PageCount Then
rL.AbsolutePage=Me.Data("@page")
Else
If rL.RecordCount>rL.PageSize Then
Else
Exit Function
End If
End If
If Me.Data("max")>Me.Data("@size")Then
Me.Data("__nav")=RSPaging(rL,IIf(IsTrue(Me.Data("gradual")),"gradual",""))
End If
If IsTrue(Me.Data("keys"))Then
Do While Exists(rL)And iCount<=rL.PageSize
Me.Data("ids")=Enclose(Me.Data("ids"),"",",")&rL(Me.Source.key)
iCount=iCount+1
rL.MoveNext
Loop
KillRS rL
iCount=1
If Me.Data("ids")="" Then DBG.FatalError "Missing IDs in list.print"
If Me.Source.KeyQuotes Then
Me.Data("ids")=rxReplace(Me.Data("ids"),"([^,]+)","'$1'")
End If
Me.Data("sql")=IIf(IsTest,"/*list.IDs-DATA("&Me.Data("datasrc")&")*/","")&"SELECT /*"&Me.Data("@top")&"*/"&IIf(Me.Data("@top")>0,"TOP "&Me.Data("@top")&" ","")&TOT(Me.Data("fields"),"%FIELDS")&" FROM %TABLE %NOLOCK %JOIN %GROUP WHERE t.%KEY IN ("&Me.Data("ids")&") "&IIf(Me.Source.Multikey," AND ("&Me.Data("criteria")&") ","")&Me.Data("sort")&""
Set rL=Me.Source.Run(Me.Data("sql"))
rL.PageSize=Me.Data("@size")
End If
Do While Exists(rL)And iCount<=rL.PageSize And dL("-ROW")<=Me.Data("rows")And iCount<=Me.Data("max")
Set parser.RS=rL
blnPrintRec=True:blnAddCount=True
If iCount<Me.Data("startFrom")Then
blnPrintRec=False
End If
If IsSet(Me.Data("skip"))Then
If IsInList(dL(Me.Data("@K")),Me.Data("skip"))Then
blnPrintRec=False
blnAddCount=False
End If
End If
If blnPrintRec Then
dL("i")=iCount
dL("#")=iCount
UpdateCollection dL,rL
If IsTrue(Me.Data("single"))Then
Else
If(iCount=1)Or((iCount-1)Mod Me.Data("cells"))=0 Then
dL("-ROW")=VOZ(dL("-ROW"))+1
dL("-RID")=dL("-ROW")
dL("-REven")=IsTrue(CBool((dL("-ROW")Mod 2)=0))
If Me.Data("nodiv")="" Then sListQS.Add " <div class=""Row Row"&dL("-ROW")&" R"&IIf(dL("-REven"),"Even","Odd")&" Clear"">"
End If
dL("-CEL")=VOZ(dL("-CEL"))+1
If dL("-CEL")>Me.Data("cells")Then dL("-CEL")=1
sID=dL(Me.Data("@K"))&""
dL("-CID")=dL("-CEL")
If VOZ(Me.Data("cells"))=1 Then
dL("-CEven")=dL("-REven")
Else
dL("-CEven")=IsTrue(CBool((dL("-CEL")Mod 2)=0))
End If
bVF=IIf(dL("i")=1,"y","")
bVL=IIf(dL("i")=dL("max")Or dL("i")=Me.Data("@last"),"y","")
If dL("cells")>1 Then
bRF=IIf(dL("-CEL")=1,"y","")
bRL=IIf(dL("-CEL")=dL("cells"),"y","")
Else
bRF=bVF
bRL=bVL
End If
If Me.Data("nodiv")="" Then
sListQS.Add "<div class=""Item "'Item" & TOT(sID,iCount) & " "
sListQS.Add "C"&dL("-CEL")&" C"&IIf(dL("-CEven"),"Even","Odd")&" "
sListQS.Add "R"&dL("-ROW")&" R"&IIf(dL("-REven"),"Even","Odd")&" "
sListQS.Add IIf(TOT(bVF,bRF),"first not-last","not-first")&" "
sListQS.Add IIf(TOT(bVL,bRL),"not-first last","not-last")&" "
sListQS.Add """>"
End If
End If
dL("Reven")=IIf(dL("-REven"),"y","n")
dL("Rodd")=IIf(dL("-REven"),"n","y")
dL("Ceven")=IIf(dL("-CEven"),"y","n")
dL("Codd")=IIf(dL("-CEven"),"n","y")
dL("reven")=dL("Reven")
dL("rodd")=dL("Rodd")
dL("ceven")=dL("Ceven")
dL("codd")=dL("Codd")
dL("Rfirst")=IIf(bRF,"y","n")
dL("Rlast")=IIf(bRL,"y","n")
dL("-first")=IIf(bVF,"y","n")
dL("-last")=IIf(bVL,"y","n")
dL("-row")=VOZ(dL("-ROW"))
dL("-cell")=VOZ(dL("-CEL"))
dL("-odd")=IIf(dL("Codd"),"y","n")
dL("-even")=IIf(dL("Ceven"),"y","n")
Dim rT,fL,fN,fV,row_name
If Me.TData("stats")<>"" And IsObject(Me.Data("rT"))Then
row_name=TOT(RSFieldGet(rL,"name"),"Unknown")
Set rT=Me.Data("rT")
For Each fL In Split(rT("stats-fields"),",")
fN=fL
fV=rT(fN&"-"&"sum")
If fV<>"" Then
fV=VOZ(fV)
dL(fN&"-"&"share")=FormatNumber(VOZ(dL(fN))/ fV,4)
If VOZ(dL(fN&"-"&"share"))>VOZ(Me.Data("min-share"))Then
rT(fN&"-"&"chart-values")=Enclose(rT(fN&"-"&"chart-values"),"",",")&dL(fN&"-"&"sum")
rT(fN&"-"&"chart-shares")=Enclose(rT(fN&"-"&"chart-shares"),"",",")&FormatNumber(dL(fN&"-"&"share")*100,1)
rT(fN&"-"&"chart-labels")=Enclose(rT(fN&"-"&"chart-labels"),""," |")&URLEncode(Replace(TOE(row_name),"|",""))
rT(fN&"-"&"chart-labels-values")=Enclose(rT(fN&"-"&"chart-labels-values"),""," |")&Replace(FormatNumber(VOZ(RSFieldGet(rL,fN)),2),"|","")
Else
End If
End If
Next
UpdateCollection dL,rT
Set Me.Data("rT")=rT
End If
Dim ttt
ttt=tag__DATA_MergeWithObject(dL)
sListQS.Add ttt
If IsTrue(Me.Data("single"))Then
Else
If Me.Data("nodiv")="" Then sListQS.Add "</div>"
If(iCount Mod Me.Data("cells"))=0 Then
If Me.Data("nodiv")="" Then sListQS.Add "</div>"
bRowOpen=False
Else
bRowOpen=True
End If
End If
End If
If blnAddCount Then
iCount=iCount+1
End If
If IsSet(Me.Data("random"))Then
rL.MoveFirst
Else
rL.MoveNext
End If
Loop
KillRS rL
If bRowOpen Then
If Me.Data("nodiv")="" Then sListQS.Add "</div>"
End If
If Me.TData("stats")<>"" And IsObject(Me.Data("rT"))Then
Set rT=Me.Data("rT")
rT("TOTAL")=True
rT("name")=TOT(Me.TData("TOTAL-name,stats-name"),"TOTAL")
For Each fL In Split(rT("stats-fields"),",")
fN=fL
rT(fN)=rT(fN&"-"&"sum")
rT(fN&"-"&"share")=1
Next
UpdateCollection dL,rT
sListQS.Add tag__DATA_MergeWithObject(dL)
Set rT=Nothing
End If
Dim o
o=sListQS
If o="" Then
Else
Me.Data("list-head")=Me.TData("before-list,list-before,list-head,list-header")
If Me.Data("list-head")<>"" Then
o=MergeTemplate(Me.Data("list-head"),dL)&o
End If
Me.Data("list-foot")=Me.TData("after-list,list-after,list-foot,list-footer")
If Me.Data("list-foot")<>"" Then
o=o&MergeTemplate(Me.Data("list-foot"),dL)
End If
o=o&Me.Data("__nav")
If Me.Data("nodiv")="" And Query.FQ("RSPaging")="" Then
o=Enclose(o,"<div class="""&TOT(Me.Data("class"),"")&""">"&vbCRLf,vbCRLf&"</div>")
End If
tag_list__PRINT=o
End If
End Function
Public Function tag_list():tag_list=True
Call tag__DATA()
Dim rL
Dim iii,ttt,xxx
If IsObject(Me.Data("rs"))Then
Set rL=Me.Data("rs")
Else
If Me.Data("rs")<>"" Then
If Me.Data("rs")="!" Then
Set rL=j("rs")
Else
Set rL=OBJ2(Me.Data("rs"))
End If
End If
End If
If IsObject(rL)Then
Else
If IsObject(Me.Source)Then
If Me.Source Is Nothing Then Exit Function
Else
Exit Function
End If
End If
If(Not IsObject(rL))Then
If Me.Source.MultiKey Or Me.Source.DirectDataMode Then Me.Data("direct")=True
Me.Data("filter")=MergeCriteria(Me.Data("filter"),"AND",Me.Data("root-filter"))
If Me.Data("sql")<>"" Then
Me.Data("-sql")=Me.Source.ParseSQL("/*list.sql*/"&Me.Data("sql")&"")
Set rL=Me.Source.Run(Me.Data("-sql"))
If Me.Data("saveCountIn")<>"" Then j(Me.Data("saveCountIn"))=rL.RecordCount
ElseIf Me.Data("group")<>"" Or Me.Source.Group<>"" Then
If Me.Data("group")<>"" Then
Me.Data("group")="GROUP BY "&rxReplace(rxReplace(Me.Data("group"),"\s*GROUP\s*BY\s*",""),"\s*AS\s+[^,]+","")
End If
Me.Data("-sql")=Me.Source.ParseSQL("/*list.group*/SELECT "&IIf(GTZ(Me.Data("max")),"TOP "&VOZ(Me.Data("max"))&" ","")&TOT(Me.Data("fields"),"%FIELDS")&" FROM %TABLE %NOLOCK %JOIN WHERE "&Me.Data("filter")&" "&TOT(Me.Data("group"),"%GROUP")&" "&Me.Data("sort")&" ")
Me.Data("-sql")=Replace(Me.Data("-sql"),"t. convert(","convert(")
Me.Data("-sql")=Replace(Me.Data("-sql"),"t.convert(","convert(")
Set rL=Me.Source.Run(Me.Data("-sql"))
If Exists(rL)Then
If Me.Data("saveCountIn")<>"" Then j(Me.Data("saveCountIn"))=rL.RecordCount
End If
Else
Me.Data("-sql")=Me.Data("--sql")
If VOZ(Me.Data("max"))<0 Then Me.Data("direct")="y"
If VOZ(Me.Data("recpp"))<0 Then Me.Data("direct")="y"
If Me.Data("mode")="direct" Then Me.Data("direct")="y"
If IsTrue(Me.Data("direct"))Then
Me.Data("-sql")="/*list.DATA*/"&Me.Data("-sql")
If GTZ(Me.Data("@size"))Then
If rxTest(Me.Data("-sql"),"SELECT\s+TOP")Then
Else
Me.Data("-sql")=Replace(Me.Data("-sql"),"%FIELDS","TOP "&VOZ((VOV(Me.Data("@size"),FWX_DEFAULT_RECPP)*Me.Data("@page"))+1)&" %FIELDS")
Me.Data("gradual")="y"
End If
End If
Me.Data("-sql")=Replace(Me.Data("-sql"),"%FIELDS",TOT(Me.Data("fields"),"%FIELDS")&"")
Else
Me.Data("-sql")="/*list.IDs*/"&Me.Data("-sql")
If Me.Data("@size")<=0 Or Me.Data("max")>0 Or Me.Data("@page")>5 Then
Me.Data("-sql")=Replace(Me.Data("-sql"),"%FIELDS","t.%KEY")
Else
Me.Data("-sql")=Replace(Me.Data("-sql"),"%FIELDS","TOP "&VOZ((VOV(Me.Data("@size"),FWX_DEFAULT_RECPP)*Me.Data("@page"))+1)&" t.%KEY")
Me.Data("gradual")="y"
End If
Me.Data("keys")=True
End If
Set rL=Me.Source.Run(Me.Data("-sql"))
If Exists(rL)Then
If Me.Data("saveCountIn")<>"" Then j(Me.Data("saveCountIn"))=rL.RecordCount
End If
End If
End If
Set j("rs")=rL
Dim rT,fL,fN,fV
Me.Data("stats")=Me.TData("stats,aggregate")
If Me.TData("stats")<>"" Then
If Exists(rL)Then
Set rT=Dictionary()
Do While Exists(rL)
Me.Data("min-share")=VOV(Me.NData("chart-threshold,min-share,share-threshold"),0.1)
rT("stats-count")=VOZ(rT("count"))+1
rT("row"&"-"&"labels")=Enclose(rT("row"&"-"&"labels"),""," |")&URLEncode(Replace(TOE(RSFieldGet(rL,"name")),",",""))
For Each fL In Split(Me.Data("stats"),",")
fN=TOE(fL)
fV=RSFieldGet(rL,fN)
If fV<>"" Then
fV=VOZ(fV)
If rT("stats-fields-"&fN)="" Then
rT("stats-fields-"&fN)=True
rT("stats-fields")=Enclose(rT("stats-fields"),"",",")&fN
End If
rT(fN&"-"&"values")=Enclose(rT(fN&"-"&"values"),"",",")&fV
rT(fN&"-"&"labels")=Enclose(rT(fN&"-"&"labels"),""," |")&URLEncode(Replace(TOE(RSFieldGet(rL,"name")),"|",""))
rT(fN&"-"&"sum")=VOZ(rT(fN&"-"&"sum"))+fV
If GTZ(rT("quantity"&"-"&"sum"))Then
rT(fN&"-"&"avg")=VOZ(rT(fN&"-"&"sum"))/ VOZ(rT("quantity"&"-"&"sum"))
Else
rT(fN&"-"&"avg")=VOZ(rT(fN&"-"&"sum"))/ rT("stats-count")
End If
rT(fN&"-"&"avg")=FormatNumber(rT(fN&"-"&"avg"),4)
If rT(fN&"-"&"min")="" Or fV<VOZ(rT(fN&"-"&"min"))Then rT(fN&"-"&"min")=fV
If rT(fN&"-"&"max")="" Or fV>VOZ(rT(fN&"-"&"max"))Then rT(fN&"-"&"max")=fV
If rT(fN&"-"&"first")="" Then rT(fN&"-"&"first")=fV
rT(fN&"-"&"last")=fV
Else
End If
Next
rL.MoveNext()
Loop
rL.Movefirst()
Set Me.Data("rT")=rT
End If
End If
If Not Exists(rL)Then
If(IsLocal And Me.Data("debug")<>"")Then
Me.Data("_output")="LIST NOT FOUND<br/>SRC:"&Replace(rL.Source,"[","[?")&"<br/>SQL:"&Replace(Me.Data("sql"),"[","[?")&""
End If
Else
j("x-TAG-LIST")=VOZ(j("x-TAG-LIST"))+1
iii=rL.RecordCount
ttt=LCase(Me.Data("template"))
Select Case ttt
Case "doaccount":xxx=DoAccount(Me.Source.Run(Me.Data("-sql")),Me.Data("format"))
Case "doaddress":xxx=DoAddress(Me.Source.Run(Me.Data("-sql")),Me.Data("format"))
Case "doitems":xxx=DoItems(Me.Source.Run(Me.Data("-sql")),Me.Data("format"))
Case "dojobs":xxx=DoJobs(Me.Source.Run(Me.Data("-sql")),Me.Data("format"))
Case "domodifiers":xxx=DoModifiers(Me.Source.Run(Me.Data("-sql")),Me.Data("format"))
Case Else:xxx=tag_list__PRINT(rL)
End Select
Me.Data("_output")=xxx&""
End If
KillRS j("rs"):j("rs")=""
KillRS rL
End Function
Public Function tag_menu__NODE__FILTER(ByRef rs,ByVal sF)
Err.Clear():On Error Resume Next
rs.Filter=sF
If Err.Description<>"" Then
DBG.Warning "INVALID FILTER: "&sF
End If
Err.Clear()
End Function
Public Function tag_menu__NODE__SORT(ByRef rs)
Err.Clear():On Error Resume Next
rs.Sort=TOT(Me.Data("sort"),"sort ASC, name ASC")
If Err.Description<>"" Or Err.Number<>0 Then:Err.Clear():rs.Sort="sort ASC"
If Err.Description<>"" Or Err.Number<>0 Then:Err.Clear():rs.Sort="name ASC"
If Err.Description<>"" Or Err.Number<>0 Then:Err.Clear():rs.Sort=Me.Data("@K")&" DESC"
If Err.Description<>"" Or Err.Number<>0 Then
DBG "sorting error"
rs.Sort=""
End If
Err.Clear()
End Function
Public Function tag_menu__NODE(ByVal rMD,ByVal iRoot,ByVal iLayer,ByVal dCurItem)
Dim output:Set output=New QuickString
If Not IsObject(rMD)Then Exit Function
If iLayer>Me.Data("layers")And Me.Data("layers")>0 Then Exit Function
output.Add MergeTemplate(Me.Data("submenu-template"),dCurItem)
If IsTrue(Me.Data("recurs"))Then
Set rsMenuItems=rMD.Clone()
Dim sF:sF="[parent] = "&iRoot&""
tag_menu__NODE__FILTER rsMenuItems,sF
If Exists(rsMenuItems)Then
If Me.Data("sorted")="" Then
tag_menu__NODE__SORT rsMenuItems
End If
Me.Data("menu-i")=0
Me.Data("menu-n")=rsMenuItems.RecordCount
Do While Exists(rsMenuItems)
Me.Data("menu-i")=Me.Data("menu-i")+1
output.Add tag_menu__NODE_ITEM(rMD,rsMenuItems,iLayer)
rsMenuItems.MoveNext
Loop
KillRS rsMenuItems
End If
sFilter=""
End If
sThisMenuID="":sClassName="":sMenuProperties=""
Dim iThisLayer:iThisLayer=(iLayer+Me.Data("starting-layer"))
sClassName="l"&iThisLayer&" "
If Me.Data("root")=iRoot And(Not IsTrue(Me.Data("isSubmenu")))Then
sThisMenuID=" id="""&Me.Data("id")&""""
sClassName=sClassName&"ListMenu "
sClassName=sClassName&Me.TData("class")&" "
Else
sThisMenuID=" id="""&Me.Data("id")&"m"&iRoot&""""
sClassName=sClassName&Me.TData("ul-class")&" "
End If
If IsTrue(Me.TData("bootstrap-dropdown"))And iThisLayer=2 Then
sClassName=sClassName&" dropdown-menu"
End If
If Me.Data("root")=iRoot And(IsTrue(Me.Data("isSubmenu")))Then
Else
If output<>"" Then
Dim ul:ul=TOT(Me.TData("ul,ulTag"),"ul")
output="<"&ul&sThisMenuID&" class="""&Trim(sClassName)&""">"&output&"</"&ul&">"
End If
End If
tag_menu__NODE=output
End Function
Public Function tag_menu__NODE_ITEM(ByVal rMD,ByRef rsMenuItem,ByVal iLayer)
Dim dL:Set dL=Dictionary()
UpdateCollection dL,Me.Data
UpdateCollection dL,rsMenuItem
dL("i")=VOZ(dL(TOT(dL("@K"),dL("k"))))
dL("l")=VOZ(iLayer+VOZ(Me.Data("starting-layer")-1))
Dim li:li=TOT(Me.TData("li,ulTag"),"li")
Dim o:Set o=New QuickString
With o
.After=vbCR
If rxTest(dL("name"),"divider?|gap|spacer?")Then
.Add "<"&li&" class=""Divider"">&nbsp;</"&li&">"
Else
If dL("a-href")="" Then dL("a-href")=Me.TData("href,link,target,url")
If dL("a-href")="" Then dL("a-href")=dL("a-link")
If dL("a-href")="" Then dL("a-href")="[tot a='|redirect|' b='|link|']"
If dL("a-id")="" Then dL("a-id")=Me.Data("id")&"a|i|"
dL("li-class")=Enclose(dL("li-class"),""," ")&"c"&dL("m")&"n"&dL("i")&" l"&dL("l")&IIf(Me.Data("menu-i")=1," first","")&IIf(Me.Data("menu-i")=Me.Data("menu-n")," last","")&""
.Add "<"&li&" id='"&Me.Data("id")&"-"&Me.Data("@K")&"-"&dL("i")&"' class='"&dL("li-class")&"'>"
dL("a-class")=Enclose(dL("a-class"),""," ")&"c"&dL("m")&"a"&dL("i")&" l"&dL("l")&IIf(Me.Data("menu-i")=1," first","")&IIf(Me.Data("menu-i")=Me.Data("menu-n")," last","")&""
If Me.TData("noLink,listOnly,no-link")="" Then
.Add "<a"
For Each sE In Split("href,id,class,title,target,tooltip",",")
If dL("a-"&sE)<>"" Then
.Add " "&sE&"="""&MergeTemplate(dL("a-"&sE),dL)&""""
End If
Next
For Each sE In Split("click,dblclick,mousedown,mouseup,mouseover,mouseout",",")
If dL("on"&sE)<>"" Then
.Add " on"&sE&"="""&MergeTemplate(dL("on"&sE),dL)&""""
End If
Next
.Add ">"
End If
.Add MergeTemplate(Me.Data("template"),dL)
If Me.TData("noLink")="" Then
.Add "</a>"
End If
If dL("i")>0 And dL("parent")<>"" And dL("location")<>"" Then
dL("parent")=VOZ(dL("parent"))
dL("location")=dL("location")&""
If dL("submenus")="" Then
.Add tag_menu__NODE(rMD,dL("i"),iLayer+1,dL)
Else
If GTZ(dL("submenus"))Then
If rxTest(dL("location"),"\b"&dL("submenus")&"\b")Then
.Add tag_menu__NODE(rMD,dL("i"),iLayer+1,dL)
End If
Else
dL("submenus")=","&rxReplace(dL("submenus"),"\W+",",")&","
If rxTest(dL("submenus"),"\b"&VOZ(dL("i"))&"\b")Then
.Add tag_menu__NODE(rMD,dL("i"),iLayer+1,dL)
End If
End If
End If
End If
.Add "</"&li&">"
End If
o=Replace(o,"-quote-","'")
End With
tag_menu__NODE_ITEM=o
End Function
Public Function tag_menu():tag_menu=True
If Not IsObject(Me.Source)Then Exit Function
Dim rMD
If Me.Data("recurs")Then Me.Data("wild")=True
Me.Data("root")=Me.NData("root")
If j("tag-menu-count")="" Then j("tag-menu-count")=0
j("tag-menu-count")=j("tag-menu-count")+1
If Me.Data("m")="" Then Me.Data("m")="m"&j("tag-menu-count")
If Me.Data("cache")="" Then
If Me.Data("kache")="e" Then Me.Data("kache")=LCase(Me.TData("datasrc,entity,ent,e")&Enclose(Me.JData("where,filter",","),"{,","}"))
If Me.Data("kache")="" Then Me.Data("kache")=Me.TagKey()
Else
Me.Data("kache")=""
End If
DBG "tag_MENU> CACHE: "&Me.Data("kache")&""
If Me.Data("kache")<>"" Then
If output="" Then
output=Application.Contents("#C-TAG-"&Me.Data("kache"))
If output<>"" Then
j("_STAT-menu-cache-ram")=VOZ(j("_STAT-menu-cache-ram"))+1
DBG "tag_MENU> CACHE FOUND (M): "&Len(output)&" bytes"
End If
End If
If output="" Then
Me.Data("fache")=Me.Data("kache")
Me.Data("fache")=MyDomainCache&"/TAGS/"&Me.ClassName&"/"&Me.Data("fache")&"/E~"&j("EDITION")&"_R~"&j("REVISION")&".txt"
Me.Data("fache")=Replace(Me.Data("fache"),"TAG#@","")
Me.Data("fache")=Replace(Me.Data("fache"),"@","/")
output=ReadFileWithoutValidation(Me.Data("fache"))
If output<>"" Then
j("_STAT-menu-cache-fso")=VOZ(j("_STAT-menu-cache-fso"))+1
DBG "tag_MENU> CACHE FOUND (F): "&Len(output)&" bytes"
End If
End If
End If
If output="" Then
j("_STAT-menu-fresh")=VOZ(j("_STAT-menu-fresh"))+1
End If
If output="" Then
Call tag__DATA()
Me.Source.Filters("_customSTAKEHOLDER")=""
Me.Data("sql")=Me.Data("--sql")
Me.Data("sql")=Replace(Me.Data("sql"),"%FIELDS",TOT(Me.Data("fields"),"%FIELDS"))
Set rMD=Me.Source.Run(Me.Data("sql"))
If Not ExistsOrKill(rMD)Then Exit Function
Me.Data("id")=TOT(Me.Data("id"),Me.Data("m"))
Me.Data("ActiveStatic")=TOT(Me.Data("ActiveStatic"),"yes")
Me.Data("isSubmenu")=Me.BData("isSubmenu,slave,sub,submenu,childrenonly,chilc,childonly")
Me.Data("maxLength")=VOV(Me.NData("max,maxLength"),25)
Me.Data("layers")=Me.NData("layers,levels,branches")
Me.Data("layers-start")=VOV(Me.NData("layer,start-layer,starting-layer,offset-layers"),1)
tag__DATA__FindTemplate "submenu-"
output=tag_menu__NODE(rMD,Me.Data("root"),1,Me.Data)
KillRS rMD
output=TOT(output,"!EMPTY!")
If Me.Data("kache")<>"" Then
Application.Contents("#C-TAG-"&Me.Data("kache"))=output
If Me.Data("fache")<>"" Then WriteToFile Me.Data("fache"),output
End If
End If
If output="!EMPTY!" Then output=""
If Me.Data("notselected")="" Then
If Me.Data("current")="" Then
Me.Data("current")=j("__PAGE")
End If
If Me.Data("current")<>"" Then
If InStr(1,Me.Data("current"),"{")>0 Then
Me.Data("location")=Me.Data("current")
Else
r.Pattern="\b(c"&Me.Data("m")&"n("&Me.Data("current")&"))\b"
output=r.Replace(output,"$1 Current")
End If
End If
If Me.Data("location")="" Then
Me.Data("location")=TOT(j("__LOCATION"),j("pageMap"))
End If
If Me.Data("location")<>"" Then
sLoc=Me.Data("location")
sLoc=Replace(sLoc,"{","(")
sLoc=Replace(sLoc,"}",")")
sLoc=Replace(sLoc,")(",")|(")
r.Pattern="\b(c"&Me.Data("m")&"(a|n)("&sLoc&"))\b"
output=r.Replace(output,"$1 Active")
End If
End If
Me.Data("_output")=output
Err.Clear()
End Function
Public Function tag_modules():tag_modules=True:With Me
.Data("template")="options"
.Data("current")=.TData("current,match,cur")
.Data("a")="/$TMPL_FOLDER/modules.json"
Dim f,x
For Each x In Split(FE__TMPL_Folders&";"&BO__TMPL_Folders,";")
If x<>"" Then
If Me.Data("p")="" Then
f=x&"modules.json"
If FileExists(f)Then
Me.Data("p")=f
End If
End If
If Me.Data("p")="" Then
f=x&"fe/modules.json"
If FileExists(f)Then
Me.Data("p")=f
End If
End If
If Me.Data("p")="" Then
f=x&"fe/@/modules.json"
If FileExists(f)Then
Me.Data("p")=f
End If
End If
If Me.Data("p")="" Then
f=x&"fe/base/modules.json"
If FileExists(f)Then
Me.Data("p")=f
End If
End If
If Me.Data("p")="" Then
f=x&"bo/modules.json"
If FileExists(f)Then
Me.Data("p")=f
End If
End If
If Me.Data("p")="" Then
f=x&"bo/@/modules.json"
If FileExists(f)Then
Me.Data("p")=f
End If
End If
End If
Next
.Data("x")="/$SITE_FOLDER/modules.json"
.Data("y")="/fwx/fe/modules.json"
.Data("z")="/fwx/bo/modules.json"
Call tag_json()
End With:End Function
Public Function tag_json():tag_json=True:With Me
Call tag__String()
Call tag__DATA_Prepare_Template()
Dim m:m=Me.Data("max")
Dim t:t=Me.Data("template")
If t="" Then DBG.FatalError "tag_xml error: template not found"
Dim o:Set o=New QuickString
o.Divider=Me.TData("divider,divide,separator,div,join")
Dim f,json,x,jObj,jSlf,jArr,node,dLoc,tag_json_merge_single_node
Me.Data("json")=Me.TData("json,url,file,data,source,datasrc,rss,feed,uri,u,string")
If Me.TData("json,string,p,q,r,x,y,z")="" Then Exit Function
json=""
For Each f In Split("json,string,a,b,c,d,e,p,q,r,x,y,z",",")
If IsObject(Me.Data(f))Then
Else
If json="" Then
f=.Data(f)
If f<>"" And .Data("empty-"&f)="" Then
If StartsWith(f,"/fwx/")Then
json=ReadFile(f)
Else
If StartsWith(f,"/")Then
f=TranslatedPath(f)
End If
json=GetResponse(f)
End If
If json="" Then
.Data("empty-"&f)="y"
End If
End If
End If
End If
Next
.Data("json")=json
If json="" Then Exit Function
.Data("node")=.TData("node,root,selector")
If .Data("node")<>"" Then
Set jObj=STR2JSON(.Data("json"))
If rxTest(node,"^(root|self)$")Then
Set jSlf=jObj
.Data("mode")="single"
ElseIf .TData("mode")="single" Then
Set jSlf=jObj(.Data("node"))
Else
jArr=jObj(.Data("node"))
End If
Set jObj=Nothing
Else
If .TData("mode")="single" Then
Set jSlf=jObj
Else
jArr=STR2Array(.Data("json"))
End If
End If
If .Data("mode")="single" Then
Set dLoc=Dictionary()
dLoc("i")=1
UpdateCollection dLoc,.Data
UpdateCollectionJSON dLoc,node
tag_json_merge_single_node=MergeTemplate(t,dLoc)
Set dLoc=Nothing:dLoc=""
o.Add tag_json_merge_single_node
Else
If IsArray(jArr)Then
Else
j("FatalError-Dump")=json
DBG.FatalError "JSON: invalid data, not an array"
End If
For Each node In jArr
If m>0 And i>=m Then
Else
i=i+1
Set dLoc=Dictionary()
dLoc("i")=i
UpdateCollection dLoc,.Data
UpdateCollectionJSON dLoc,node
tag_json_merge_single_node=MergeTemplate(t,dLoc)
Set dLoc=Nothing:dLoc=""
o.Add tag_json_merge_single_node
End If
Next
End If
Set jObj=Nothing:jObj=""
Set jArr=Nothing:jArr=""
.Data("_output")=o
End With:End Function
Public Function tag_xml():tag_xml=True
Call tag__String()
Call tag__DATA()
Me.Data("xml")=Me.TData("xml,url,data,source,datasrc,rss,feed,uri,u,string")
If Me.Data("xml")="" Then Exit Function
If Me.Data("picasa")<>"" Then
Select Case LCase(Me.Data("picasa"))
Case "albums":Me.Data("xml")=PicasaAlbums(Me.Data("xml"))
Case "photos":Me.Data("xml")=PicasaPhotos(Me.Data("xml"))
Case Else:Me.Data("xml")=PicasaPhotos(Me.Data("xml"))
End Select
End If
Me.Data("nodes")=Me.TData("nodes,node,repeat,on,for,each,row,rows,selector")
If Me.Data("max")=0 Then Me.Data("max")=FWX_DEFAULT_RECPP
Dim m:m=Me.Data("max")
Dim s:s=Me.Data("xml")
Dim n:n=Me.Data("nodes")
Dim t:t=Me.Data("template")
If t="" Then DBG.FatalError "tag_xml error: template not found"
Dim o:Set o=New QuickString
o.Divider=Me.TData("divider,divide,separator,div,join")
If Not IsTrue(Me.Data("always-fresh"))Then
Me.Data("cache.hash")=MD5(s)
Me.Data("cache.read")=Application.Contents("XML-"&Me.Data("cache.hash"))
If Me.Data("cache.read")<>"" Then
End If
Else
End If
If Me.Data("cache.read")="!empty" Then
ElseIf s<>"" Or Me.Data("cache.read")<>"" Then
Dim node,nodes,i:i=0
Dim oXML
DBG ""
If Me.Data("cache.read")<>"" Then
Set oXML=FlatXml(Me.Data("cache.read"))
Else
Set oXML=FlatXml(s)
End If
DBG ""
If oXML.xml<>"" Then
If Me.Data("cache.hash")<>"" And Me.Data("cache.read")="" Then
DBG "tag_xml: SAVE CACHE > cache.hash= "&Me.Data("cache.hash")
Application.Contents("XML-"&Me.Data("cache.hash"))=TOT(oXML.xml,"!empty")'Me.Data("cache.read")
End If
Dim xml_raw,xml_len
xml_raw=TOE(oXML.xml)
xml_len=Len(xml_raw)
If n<>"" Then
Set nodes=oXML.selectNodes(n)
Else
Set nodes=oXML.selectNodes("//entry")
If nodes.Length=0 Then
Set nodes=oXML.selectNodes("//item")
End If
End If
For Each node In nodes
If m>0 And i>=m Then
Else
i=i+1
tag_xml_merge_single_node=XmlMerge(node,t)
o.Add tag_xml_merge_single_node
End If
Next
nodes="":oXML=""
o=o&""
If o="" Then
o="<a title='XML Feed A - output not generated (xml_len="&xml_len&")' href='"&Encode(Me.Data("xml"))&"' class='Disabled'>.</a>"
End If
Else
DBG.Warning "XML LOAD FAIL "&s&""
If o="" Then
o="<a title='XML Feed B - feed not loaded' href='"&Encode(Me.Data("xml"))&"' class='Disabled'>"&IIf(IsTest,"feed not found",".")&"</a>"
End If
End If
End If
Me.Data("_output")=o
End Function
Public Function tag_quote():tag_quote=True
If Me.Data("max")="" Then Me.Data("max")=-1
If Me.Data("recpp")="" Then Me.Data("recpp")=-1
Me.Data("direct")="y":Me.Data("mode")="direct"
Me.Data("NoDiv")="y"
Me.Data("bit")=LCase(Me.TData("bit,do,piece,part,list,show,section,action"))
If Me.Data("bit")="" Then Me.Data("bit")="quote"
Me.Data("ima")="quote"
Call tag__Number()
Me.Data("quote")=Me.NData("quote,number,string,id,i")
If Not Me.Data("quote")>0 Then
Me.Data("_output")="MISSING QUOTE ID (for "&Me.Data("bit")&")"
Exit Function
End If
If rxTest(Me.Data("bit"),"(item|job|coupon|discount|adjustment)s?")Then
Me.Data("bit")=rxReplace(Me.Data("bit"),"s$","")
Me.Data("datasrc")="quot.quotes."&Me.Data("bit")&"s"
Me.Data("where")=Enclose(Me.Data("filter"),"(",") AND ")&"%FILTERS AND t.quote="&Me.Data("quote")&" AND t.active=1 /*tag_quote-a*/"
If Me.Data("TemplateFile")="" Then
Me.Data("TemplateFile")="quot/output/modifier"
If rxTest(Me.Data("bit"),"item|job")Then
Me.Data("TemplateFile")="quot/output/"&Me.Data("bit")
End If
End If
If Me.Data("fields")="" Then
Select Case LCase(Me.Data("bit"))
Case "item":Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice,discount, code, total, vartax,fixtax,status, corporate,vip, raw_gross,raw_net,raw_tax, tax, gross,cost,net,quantity,product,image"
Case "job":Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice,discount, code, total, vartax,fixtax,status, corporate,vip, raw_gross,raw_net,raw_tax, tax, gross,cost,net,date,time,reference,[option],service, sc, subs, varprice"
Case "coupon":Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice, code, varprice, reference"
Case Else:Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice, varprice, reference"
End Select
End If
CALL tag_list()
Else
If Me.Data("TemplateFile")="" Then Me.Data("TemplateFile")="quot/output/"&Me.Data("bit")&""
If Me.Data("datasrc")="" Then Me.Data("datasrc")="quot.Quotes"
If Me.Data("datasrc")="quote" Then Me.Data("datasrc")="quot.Quotes"
Me.Data("max")=1
If Me.Data("list")<>"" Then
Me.Data("datasrc")="quote"
Me.Data("where")=Enclose(Me.Data("filter"),"(",") AND ")&"%FILTERS AND t.quote="&Me.Data("quote")&" AND t.active=1 /*tag_quote-b*/"
Else
If IsObject(j("rs"))Then
Me.Data("_output")=TemplateFile(Me.Data("TemplateFile"))
Else
Me.Data("where")=Enclose(Me.Data("filter"),"(",") AND ")&"/*%FILTERS AND*/ t.quote="&Me.Data("quote")&" /*tag_quote step 8*/"
CALL tag_list()
End If
End If
End If
End Function
Public Function tag_invoice():tag_invoice=True
If Me.Data("max")="" Then Me.Data("max")=-1
If Me.Data("recpp")="" Then Me.Data("recpp")=-1
Me.Data("direct")="y":Me.Data("mode")="direct"
Me.Data("NoDiv")="y"
Me.Data("bit")=LCase(Me.TData("bit,do,piece,part,list,show,section,action"))
If Me.Data("bit")="" Then Me.Data("bit")="invoice"
Me.Data("ima")="invoice"
Me.Data("active-filter")=" AND t.active=1"
If IsBack And IsTrue(Me.Data("includeDeleted"))Then Me.Data("active-filter")=""
Call tag__Number()
Me.Data("invoice")=Me.NData("invoice,inv,number,string,id,i")
If Not Me.Data("invoice")>0 Then
Me.Data("_output")="MISSING INVOICE ID (for "&Me.Data("bit")&")"
Exit Function
End If
If rxTest(Me.Data("bit"),"breakdown|itemize|itemise")Then
Me.Data("TemplateFile")="shop/output/breakdown"
Me.Data("_output")=TemplateFile(Me.Data("TemplateFile"))
ElseIf rxTest(Me.Data("bit"),"(item|job|coupon|discount|adjustment|refund|surcharge)s?")Then
Me.Data("bit")=rxReplace(Me.Data("bit"),"s$","")
Me.Data("datasrc")="shop.AllInvoices."&Me.Data("bit")&"s"
If rxTest(Me.Data("bit"),"(item|job)s?")Then
End If
Me.Data("where")=Enclose(Me.Data("filter"),"(",") AND ")&"t.invoice="&Me.Data("invoice")&Me.Data("active-filter")&" /*tag_invoice-a tag_invoice-datasrc='"&Me.Data("datasrc")&"'*/"
If Me.Data("TemplateFile")="" Then
Me.Data("TemplateFile")="shop/output/modifier"
If rxTest(Me.Data("bit"),"item|job")Then
Me.Data("TemplateFile")="shop/output/"&Me.Data("bit")
End If
End If
If Me.Data("fields")="" Then
Select Case LCase(Me.Data("bit"))
Case "item":Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice,code, vartax,fixtax,tax, discount, raw_gross,raw_net,raw_tax, discount, total, status, corporate,vip, tax,gross,cost,net,quantity,product,image"
Case "job":Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice,code, vartax,fixtax,tax, discount, raw_gross,raw_net,raw_tax, discount, total, status, corporate,vip, tax,gross,cost,net,date,time,reference,[option],service, sc, subs, varprice"
Case "coupon":Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice,code, varprice, reference"
Case Else:Me.Data("fields")=Me.Data("ima")&",name,description,notes,amount,fixprice,      varprice, reference"
End Select
End If
CALL tag_list()
Else
If Me.Data("TemplateFile")="" Then
Me.Data("TemplateFile")="shop/output/"&Me.Data("bit")&""
End If
If Me.Data("bit")="finances" Then
Me.Data("fields")="active,credit,corporate,vip,amount,tax,net,gross,paid,balance,profit,debt,0-t.balance outstanding"
End If
If Me.Data("datasrc")="" Then Me.Data("datasrc")="shop.invoices.View"
If Me.Data("datasrc")="shop.Invoices" Then Me.Data("datasrc")="shop.Invoices.View"
Me.Data("datasrc")=Replace(Me.Data("datasrc"),"shop.Invoices","shop.AllInvoices")
Me.Data("max")=1
If Me.Data("list")<>"" Then
Me.Data("datasrc")="shop.AllInvoices.View"
Me.Data("where")=Enclose(Me.Data("filter"),"(",") AND ")&"t.invoice="&Me.Data("invoice")&Me.Data("active-filter")&" /*tag_invoice-b tag_invoice-datasrc='"&Me.Data("datasrc")&"'*/"
Else
If IsObject(j("rs"))Then
Me.Data("_output")=TemplateFile(Me.Data("TemplateFile"))
Else
Me.Data("where")=Enclose(Me.Data("filter"),"(",") AND ")&"t.invoice="&Me.Data("invoice")&Me.Data("active-filter")
CALL tag_list()
End If
End If
End If
End Function
Public Function tag_pickers():tag_pickers=True
Dim rPs,rPi,rP,rI,rJ
Dim o,n,i,k,t
Set o=New QuickString
Call tag__Number()
Me.Data("job")=Me.NData("job,number,string,id,i")
If Not Me.Data("job")>0 Then
Me.Data("_output")="MISSING JOB ID (for "&Me.Data("bit")&") - "&Encode(Me.Match)
Exit Function
End If
Me.Data("picker")=TOT(Me.Data("picker"),Query.FQ("picker"))
Me.Data("e")="shop.AllInvoices.Jobs.Pickers"
Call tag__DATA()
Set rPs=shop.AllInvoices.Jobs.Pickers.Find("t.job="&Me.Data("job")&" AND "&Me.Data("filter")&"")
If Not ExistsOrKill(rPs)Then Exit Function
t=0
rPs.MoveFirst
Do While Exists(rPs)
If Me.Data("ignore")<>"" And InList(Me.Data("ignore"),VOZ(rPs("picker")))Then
ElseIf Me.Data("pickers")<>"" And(Not InList(Me.Data("pickers"),VOZ(rPs("picker"))))Then
Else
If rPs("picker")>0 Then
Set rPi=invy.Products.Pickers.ByID(rPs("picker"))
Set rPi=RS2DIC(rPi)
Me.Data(rPs("picker")&"Name")=rPi("name")
n=0
If GTZ(rPi("per-item"))Then
k=rPs("items")/rPi("per-item")
n=n+Fix(k)+IIf(k<>Fix(k),1,0)
End If
If GTZ(rPi("per-unit"))Then
k=rPs("quantity")/rPi("per-unit")
n=n+Fix(k)+IIf(k<>Fix(k),1,0)
End If
If GTZ(rPi("per-weight"))Then
k=rPs("weight")/rPi("per-weight")
n=n+Fix(k)+IIf(k<>Fix(k),1,0)
End If
If GTZ(rPi("per-job"))Then
If n=0 Then n=1
n=n * rPi("per-job")
End If
If GTZ(rPi("per-invoice"))Then
If n=0 Then n=1
n=n * rPi("per-invoice")
End If
Else
n=1
End If
t=t+n
Me.Data(rPs("picker")&"-n")=VOZ(Me.Data(rPs("picker")&"-n"))+n
Me.Data("n")=VOZ(Me.Data("n"))+VOZ(Me.Data(rPs("picker")&"-n"))
KillRS rPi
End If
rPs.MoveNext
Loop
DBG "PICKERS>END"
If Me.BData("labels,label-count,pickers,picker-count,count")Then
DBG "ONLY COUNT REQUIRED ("&Me.TData("labels,label-count,pickers,picker-count,count,number")&")"&PrintCol(Me.Data)
Me.Data("_output")=Me.Data("n")
Exit Function
End If
Set rJ=shop.AllInvoices.Jobs.View.Find("t.job="&Me.Data("job")&"")
Set dJ=RS2DIC(rJ)
i=1
rPs.MoveFirst
Do While Exists(rPs)
If InList(Me.Data("ignore"),VOZ(rPs("picker")))Then
Else
dJ("n")=Me.Data("n")
dJ("picker")=rPs("picker")
dJ("pickerN")=Me.Data(rPs("picker")&"-n")
dJ("pickerName")=Me.Data(rPs("picker")&"Name")
dJ("item")=""
dJ("method")=""
dJ("lockQuantity")=""
dJ("itemize")=False
dJ("unitize")=False
dJ("familiarize")=False
If GTZ(Rps("picker"))Then
Set rPi=invy.Products.Pickers.ByID(rPs("picker"))
End If
If Exists(rPi)Then
If GTZ(rPi("per-item"))Then
dJ("itemize")=True
End If
If GTZ(rPi("per-unit"))Then
dJ("itemize")=True
dJ("unitize")=True
End If
End If
If dJ("itemize")Or dJ("unitize")Then
Set rI=shop.AllInvoices.Items.View.FindFields("t.item,t.quantity","t.invoice="&rPs("invoice")&" AND ISNULL(t.status,'') IN ('','picked','deliver') AND f.picker="&rPs("picker")&" AND %FILTER")
Do While Exists(rI)
If dJ("unitize")Then
dJ("method")="unitize"
q=1
Do While q<=rI("quantity")
dJ("item")=rI("item")
dJ("lockQuantity")=1'rI("item")
dJ("i")=i
If VOZ(rPs("picker"))=VOV(Me.NData("picker"),rPs("picker"))Then o.Add MergeTemplate(Me.Data("template"),dJ)
i=i+1
q=q+1
Loop
Else
dJ("method")="itemize"
dJ("item")=rI("item")
dJ("i")=i
If VOZ(rPs("picker"))=VOV(Me.NData("picker"),rPs("picker"))Then o.Add MergeTemplate(Me.Data("template"),dJ)
i=i+1
End If
rI.MoveNext()
Loop
KillRS rI
Else
dJ("familiarize")=True
dJ("method")="familiarize"
dJ("item")=""
k=1
Do While k<=dJ("pickerN")
dJ("i")=i
If VOZ(rPs("picker"))=VOV(Me.NData("picker"),rPs("picker"))Then o.Add MergeTemplate(Me.Data("template"),dJ)
i=i+1
k=k+1
Loop
End If
End If
rPs.MoveNext
Loop
KillRS rPs
Me.Data("_output")=o
End Function
Public Function tag__DATA()
If Me.Data("done-tag__DATA")<>"" Then Exit Function
Me.Data("done-tag__DATA")="y"
Call tag__DATA_Initialize
Call tag__DATA_Prepare_Template
Call tag__DATA_SQL
End Function
Public Function tag_CalcDate():With Me
Dim s
s=.Data("args")
s=rxReplace(s,"[^0-9\(\)\-\+\,\""a-zA-Z]","")
If GTZ(s)Then
s=DateAdd("d",Now(),VOZ(s))
Else
s=EXE2(s)
End If
If s<>"" Then
.Data("_output")=DoDate(s)
End If
End With:End Function
Public Function tag_date():tag_date=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
Call tag__iif()
.Data("date")=DoDate(.DData("date"))
.Data("is")=.TData("is,match,equals")
If .Data("is")<>"" Then
.Data("cond")=CBool(DoDate(.DData("date"))=DoDate(.DData("is")))
Else
.Data("lt")=.DData("lt,less,before")
.Data("gt")=.DData("gt,more,after")
If .TData("lt,gt")<>"" Then
Dim df
For Each df In Split("lt,gt",",")
If .Data(df)="" Then
.Data(df)="y"
Else
If df="gt" Then
.Data(df)=CBool(CDate(.Data("date"))>CDate(DoDate(.Data(df))))
Else
.Data(df)=CBool(CDate(.Data("date"))<CDate(DoDate(.Data(df))))
End If
End If
Next
.Data("cond")=CBool(IsTrue(.Data("lt"))And IsTrue(.Data("gt")))
End If
End If
If .Data("cond")<>"" Then
Call tag_iif()
.Data("date")=.Data("_output")
End if
If .Data("format")<>"" Then
.Data("date")=FmtDate(.Data("date"),.Data("format"))
End if
.Data("_output")=.Data("date")
End With:End Function
Public Function tag_dodate():tag_dodate=tag_date:End Function
Public Function tag_since():tag_since=True:With Me
If Not tag__DateTime()Then Exit Function
.Data("_output")=TimeSince(.Data("date"))
End With:End Function
Public Function tag_timesince():tag_timesince=tag_since:End Function
Public Function tag_time():tag_time=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Hour()
.Data("_output")=DoTime(.Data("date"))
End With:End Function
Public Function tag_dotime():tag_dotime=tag_time:End Function
Public Function tag_day():tag_day=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
.Data("_output")=Day(.Data("date"))
End With:End Function
Public Function tag_dayordinal():tag_dayordinal=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
.Data("_output")=FmtDayOrdinal(.Data("date"))
End With:End Function
Public Function tag_month():tag_month=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Month()
.Data("_output")=Month(.Data("date"))
End With:End Function
Public Function tag_year():tag_year=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Year()
.Data("_output")=Year(.Data("date"))
End With:End Function
Public Function tag_yyyy():tag_yyyy=tag_year:End Function
Public Function tag_shortyear():tag_shortyear=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Year()
.Data("_output")=Right(Year(DateFromString(.Data("date"))),2)
End With:End Function
Public Function tag_yy():tag_yy=tag_shortyear:End Function
Public Function tag_week():tag_week=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Week()
.Data("_output")=Week(.Data("date"))
End With:End Function
Public Function tag_wk():tag_wk=tag_week:End Function
Public Function tag_dayname():tag_dayname=True:With Me
If .Data("day")="" Then
If tag__DateTime()Then
Call tag__Date_Offset_Date()
.Data("day")=WeekDay(.Data("date"))
Else
Call tag__Number()
.Data("day")=VOZ(.Data("string"))
If .Data("day")>0 Then
Else
Exit Function
End If
End If
End If
.Data("_output")=WeekDayName(.Data("day"))
End With:End Function
Public Function tag_weekdayname():tag_weekdayname=tag_dayname:End Function
Public Function tag_wkdayname():tag_wkdayname=tag_dayname:End Function
Public Function tag_weekday():tag_weekday=tag_dayname:End Function
Public Function tag_wkday():tag_wkday=tag_dayname:End Function
Public Function tag_dname():tag_dname=tag_dayname:End Function
Public Function tag_wday():tag_wday=tag_dayname:End Function
Public Function tag_monthname():tag_monthname=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Month()
.Data("_output")=MonthName(Month(.Data("date")))
End With:End Function
Public Function tag_mname():tag_mname=tag_monthname:End Function
Public Function tag_dateandtime():tag_dateandtime=True:With Me
If Not tag__DateTime()Then Exit Function
.Data("_output")=PrintDateAndTime(.Data("date"))
End With:End Function
Public Function tag_now():tag_now=tag_dateandtime:End Function
Public Function tag_present():tag_present=tag_dateandtime:End Function
Public Function tag_current():tag_current=tag_dateandtime:End Function
Public Function tag_isToday():tag_isToday=True:Call tag__iif():With Me
Call tag__String()
.Data("b")=.TData("a,b,cond,c,d,e,string")
.Data("a")=Date()
Call tag_isday()
End With:End Function
Public Function tag_sameday():tag_sameday=tag_isday:End Function
Public Function tag_isday():tag_isday=True:Call tag__iif():With Me
If DoDate(.TData("a,a1,a2,a3,date,cond,string"))=DoDate(TOT(.TData("b,b1,b2,b3,compare,as"),Date()))Then
.Data("cond")="y"
Else
.Data("cond")="n"
End If
Call tag_iif()
End With:End Function
Public Function tag_fmtdate():tag_fmtdate=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
If .Data("format")="dynamic" Then
.Data("format")="%b '%y"
.Data("gone")=TOT(.Data("gone"),"%b-%D")
If .BData("noTime,dateOnly")Then
.Data("present")=.Data("gone")
Else
.Data("present")=TOT(.TData("present"),"<strong>%a</strong> %H:%N")
End If
End If
If Abs(DateDiff("d",Date(),.Data("date")))<3 Then
.Data("format")=TOT(.TData("sameDayFormat,sameDay,sameday,present"),.TData("format"))
ElseIf .Data("date")>Date()Then
If Abs(DateDiff("m",Date(),.Data("date")))>6 Then
.Data("format")=TOT(.TData("futureFormat,futurefmt,future"),.TData("format"))
Else
.Data("format")=TOT(.TData("soonFormat,soonfmt,soon,coming"),.TData("format"))
End If
Else
If Abs(DateDiff("m",Date(),.Data("date")))>6 Then
.Data("format")=TOT(.TData("pastFormat,pastfmt,past"),.TData("format"))
Else
.Data("format")=TOT(.TData("goneFormat,gonefmt,gone"),.TData("format"))
End If
End If
Select Case LCase(TOT(.TData("format"),.Data("args")))
Case "rfc":.Data("format")=fwxDateRFC
Case "iso":.Data("format")=fwxDateISO
Case "rss":.Data("format")=fwxDateRSS
Case "xml":.Data("format")=fwxDateXML
Case "rdf":.Data("format")=fwxDateRDF
Case "ics":.Data("format")=fwxDateICS
Case "cal":.Data("format")=fwxDateCAL
Case "ymd":.Data("format")="%y%M%D"
Case "hhnn":.Data("format")="%H%N"
Case "yymmdd":.Data("format")="%y%M%D"
Case "yyyymmdd":.Data("format")="%Y%M%D"
Case "yymmddhh":.Data("format")="%y%M%D%H"
Case "yyyymmddhh":.Data("format")="%Y%M%D%H"
Case "yymmddhhnn":.Data("format")="%y%M%D%H%N"
Case "yyyymmddhhnn":.Data("format")="%Y%M%D%H%N"
Case "timestamp":.Data("format")="%Y%M%D-%H%N"
Case "shorttimestamp":.Data("format")="%M%D-%H%N"
Case "full","big","long","longdate"
.Data("format")="%Y-%M-%D %H:%N:%S"
Case "small","short","shortdate"
.Data("format")="%d-%b-%y"
Case "time"
.Data("format")="%H:%N"
Case "weekday","wkday","wk","wd"
.Data("format")="%a"
Case Else
End Select
.Data("_output")=FmtDate(.Data("date"),.Data("format"))
If .Data("_output")="" Then .Data("_output")=DateFromString(.Data("date"))
End With:End Function
Public Function tag_yyyymmdd():tag_yyyymmdd=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
.Data("format")="yyyymmdd"
Call tag_fmtdate()
End With:End Function
Public Function tag_hhnn():tag_hhnn=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
.Data("format")="hhnn"
Call tag_fmtdate()
End With:End Function
Public Function tag_timestamp():tag_timestamp=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
.Data("format")="timestamp"
Call tag_fmtdate()
End With:End Function
Public Function tag_shorttimestamp():tag_shorttimestamp=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
.Data("format")="%M%D_%H%N"
Call tag_fmtdate()
End With:End Function
Public Function tag_fulldate():tag_fulldate=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Date()
.Data("_output")=FmtDate(.Data("date"),"%d %B %Y")
End With:End Function
Public Function tag_24hrtime():tag_24hrtime=True:With Me
If Not tag__DateTime()Then Exit Function
Call tag__Date_Offset_Hour()
.Data("_output")=FmtDate(.Data("date"),"%H:%N")
End With:End Function
Public Function tag_24hr():tag_24hr=tag_24hrtime:End Function
Public Function tag_militarytime():tag_militarytime=tag_24hrtime:End Function
Public Function tag__Calendar():With Me
Call tag__Date()
Call tag__Date_Offset_Date()
If .Data("date")<>0 Then
DIE DatePart("weekday",.Data("date"))
End If
End With:End Function
Public Function tag_compare():With Me
Dim s,d,o,v,i,x,y,z,n,p
ds="on,date,from,start,until,end"
d=Split(ds,",")
v=-1
i=0:Do While i<=UBound(d)
s=d(i):i=i+1
.Data(s)=DateFromString(TOT(.DData(s),Query.FQ(s)))
Loop
x=.TData(ds)
.Data("n")=VOZ(.TData("n,quantity,qty"))
.Data("to")=LCase(.TData("to,mode,with"))
If IsTrue(.TData("forward,ahead,future,fwd"))Then
v=1
End If
Select Case .Data("to")
Case "year","yyyy","yy","y":o=365
Case "week","wk","w":o=7
Case "month","mm","m":o=31
Case "day","dd","d":o=1
Case Else:
o=VOZ(.TData("to,offset,by,days,shift"))
v=1
End Select
If VOZ(o)<>0 Then
i=0:Do While i<=UBound(d)
s=d(i):i=i+1
If IsDate(.Data(s))Then
.Data(s)=DateAdd("d",o*v*VOV(.Data("n"),1),.Data(s))
End If
Loop
End If
If(Abs(o)>=180 Or IsTrue(.TData("adjust,relative,smart")))And IsFalse(.TData("ignore-weekday,simple,absolute"))Then
y=.TData(ds)
If IsDate(x)And IsDate(y)Then
w1=WeekDay(x)
w2=WeekDay(y)
z=IIf(w1>w2,w1-w2,w2-w1)
If z<-3 Then z=(7+z)*-1
If z>3 Then z=(z-7)*-1
End If
End If
.Data("offset")=o
.Data("vector")=v
.Data("w1")=w1
.Data("w2")=w2
.Data("x")=x
.Data("y")=y
.Data("z")=z
If VOZ(z)<>0 Then
i=0:Do While i<=UBound(d)
s=d(i):i=i+1
If IsDate(.Data(s))Then
.Data(s)=FmtDate(DateAdd("d",z,.Data(s)),"%Y-%M-%D")
End If
Loop
End If
.Data("_output")=ReAssign("~DATES&on="&.TData("date,on")&"&date="&.TData("date,on")&"&from="&.TData("from,start")&"&start="&.TData("from,start")&"&until="&.TData("until,end")&"&end="&.TData("until,end")&"&"&.TData("vars,reassign,a")&"")
End With:End Function
Public Function tag_LocalDateTime():With Me
.Data("a")=UTCNow()
.Data("timezone")=app.CFG("timezone")
Call tag_FmtDate()
End With:End Function
Public Function tag__DateTime():With Me
Call tag__String()
Call tag__Date():.Data("date")=Me.DData("date,string")
If .Data("date")="" Then
If .Data("noDefault")<>"" Then
Else
.Data("date")=Me.DData("def,default")
If .Data("date")="" And IsEmpty(.data("args"))Then
.Data("date")=UTCNow()
End If
End If
End If
If VOZ(.Data("args"))<>0 Then
.Data("offset")=VOZ(.Data("args"))
Else
.Data("offset")=.NData("offset,shift,add,args")
End If
If .Data("offset")<>0 And .Data("offset")<>"" And .Data("date")="" Then
.Data("date")=TOT(.Data("date"),Now())
End If
If IsDate(.Data("date"))Then
Else
If .Data("date")<>"" Then
DBG.FatalError "NOT A DATE:"&.Data("date")
.Data("date")=MakeDateFromString(.Data("date"))
End If
End If
If .Data("timezone")="" Or .Data("timezone")="auto" Then
If j("default-timezone")="" Then
j("default-timezone")=TOT(app.CFG("timezone")," ")
.Data("default-timezone")=j("default-timezone")
End If
.Data("timezone")=j("default-timezone")
End If
If TOE(.Data("timezone"))<>"" Then
.Data("date")=ConvertDate(.Data("date"),.Data("timezone"))
.Data("NOW")=ConvertDate(Now(),.Data("timezone"))
Else
End If
If IsDate(.Data("date"))Then
tag__DateTime=True
Else
tag__DateTime=False
.Data("date")=""
End If
End With:End Function
Public Function tag__Date_Offset_Hour():With Me
If .Data("offset")<>0 Then
.Data("date")=DateAdd("h",.Data("offset"),.Data("date"))
End If
End With:End Function
Public Function tag__Date_Offset_Day():tag__Date_Offset_Day=tag__Date_Offset_Date:End Function
Public Function tag__Date_Offset_Date():With Me
If .Data("offset")<>0 Then
.Data("date")=DateAdd("d",.Data("offset"),.Data("date"))
End If
End With:End Function
Public Function tag__Date_Offset_Week():With Me
If .Data("offset")<>0 Then
.Data("date")=DateAdd("ww",.Data("offset"),.Data("date"))
End If
End With:End Function
Public Function tag__Date_Offset_Month():With Me
If .Data("offset")<>0 Then
.Data("date")=DateAdd("mm",.Data("offset"),.Data("date"))
End If
End With:End Function
Public Function tag__Date_Offset_Year():With Me
If .Data("offset")<>0 Then
.Data("date")=DateAdd("yyyy",.Data("offset"),.Data("date"))
End If
End With:End Function
Public Function tag_LazyLoader():tag_LazyLoader=True:With Me
.Data("_output")=ReadFile("/fwx/inc/lazy/lazy.min.js")
End With:End Function
Public Function tag_site():tag_site=True:With Me
.Data("_output")="http://"&MyWWWDomain&"/"
End With:End Function
Public Function tag_home():tag_home=tag_site:End Function
Public Function tag_domain():tag_domain=tag_server:End Function
Public Function tag_server():tag_server=True:With Me
If .Match<>"[domain]" Then
Call tag__String()
.Data("_output")=rxReplace(rxSelect(.TData("string,args"),rxDomain),"^www\.","")
Else
.Data("_output")=MyDomain
End If
End With:End Function
Public Function tag_url():tag_url=True:With Me
.Data("_output")="http://"&MyWWWDomain&"/"&PAGE&""
If j("link")<>"" Then
.Data("_output")="http://"&MyWWWDomain&j("link")&""
ElseIf j("SCRIPT")<>"" Then
.Data("_output")="http://"&MyWWWDomain&j("SCRIPT")&""
End If
End With:End Function
Public Function tag_ipaddress():tag_ipaddress=tag_ip:End Function
Public Function tag_ip():tag_ip=True:With Me
.Data("_output")=CLIENT_IP
End With:End Function
Public Function tag_account():tag_account=True:With Me
.Data("_output")=MyAccount
End With:End Function
Public Function tag_username():tag_username=tag_user:End Function
Public Function tag_userid():tag_userid=tag_user:End Function
Public Function tag_user():tag_user=True:With Me
.Data("_output")=MyEmail
End With:End Function
Public Function tag_name():tag_name=True:With Me
.Data("_output")=MyName
End With:End Function
Public Function tag__GEO():tag__GEO=True:Call tag__String():With Me
Me.Data("ip")=TOT(Me.Data("string"),VISITOR_IP)
End With:End Function
Public Function tag_Geo():tag_Geo=True:Call tag__GEO():With Me
.Data("_output")=GeoIPLoc(.Data("ip"))
End With:End Function
Public Function tag_GeoData():tag_GeoData=tag_Geo:End Function
Public Function tag_CountryCode():tag_CountryCode=True:Call tag__GEO():With Me
.Data("_output")=GeoIPLocCode(.Data("ip"))
End With:End Function
Public Function tag_CountryName():tag_CountryName=True:Call tag__GEO():With Me
.Data("_output")=GeoIPLocName(.Data("ip"))
End With:End Function
Public Function tag_Country():tag_Country=tag_CountryName:End Function
Public Function tag_log():tag_log=tag_action:End Function
Public Function tag_action():tag_action=True
Dim x,n,i
n=TOT(Me.TData("name,label,description"),"")
Set x=Dictionary()
UpdateCollection x,j
UpdateCollection x,Me.Data
i=App.Actions.Log(n,x)
DBG "ACTION>LOG> (from tag)>"&i
Set x=Nothing
End Function
Public Function tag_hasAccess():tag_hasAccess=True:Call tag__String():With Me
.Data("credentials")=.TData("credentials,credential,creds,cred,have,given,string")
.Data("requirement")=.TData("requirement,require,requires,required,needs,need")
.Data("cond")=HasAccess(.Data("credentials"),.Data("requirement"))
Call tag_iif()
End With:End Function
Public Function tag_u():tag_u=True:Call tag__String():With Me
.Data("_output")=UCase(.Data("string"))
End With:End Function
Public Function tag_ucase():tag_ucase=tag_u:End Function
Public Function tag_uppercase():tag_uppercase=tag_u:End Function
Public Function tag_upper():tag_upper=tag_u:End Function
Public Function tag_l():tag_l=True:Call tag__String():With Me
.Data("_output")=LCase(.Data("string"))
End With:End Function
Public Function tag_lcase():tag_lcase=tag_l:End Function
Public Function tag_lowercase():tag_lowercase=tag_l:End Function
Public Function tag_lower():tag_lower=tag_l:End Function
Public Function tag_p():tag_p=True:Call tag__String():With Me
.Data("_output")=PCase(.Data("string"))
End With:End Function
Public Function tag_pcase():tag_pcase=tag_p:End Function
Public Function tag_propercase():tag_propercase=tag_p:End Function
Public Function tag_proper():tag_proper=tag_p:End Function
Public Function tag_q():tag_q=True:Call tag__String():With Me
If .Data("string")<>"" Then
.Data("_output")=Encode("'"&.Data("string")&"'")
End If
End With:End Function
Public Function tag_b():tag_b=True:Call tag__String():With Me
If .Data("string")<>"" Then
.Data("_output")=Encode("("&.Data("string")&")")
End If
End With:End Function
Public Function tag_brackets():tag_brackets=tag_b:End Function
Public Function tag_enclose():tag_enclose=True:Call tag__String():With Me
If .Data("string")<>"" Then
If .Data("with")<>"" Then
.Data("before")=.Data("with")
.Data("after")=.Data("with")
End If
.Data("_output")=.TData("before")&.Data("string")&.TData("after")&""
End If
End With:End Function
Public Function tag_surround():tag_surround=tag_enclose:End Function
Public Function tag_nosymbols():tag_nosymbols=True:Call tag__String():With Me
.Data("_output")=RemoveSymbols(.Data("string")&"")
End With:End Function
Public Function tag_removesymbols():tag_removesymbols=tag_nosymbols:End Function
Public Function tag_reverse():tag_reverse=True:Call tag__String():With Me
.Data("_output")=Reverse(.Data("string")&"")
End With:End Function
Public Function tag_invert():tag_invert=tag_reverse:End Function
Public Function tag_curdecimals():tag_curdecimals=True:Call tag__String():With Me
Dim c,d
c=.TData("symbol,code,country,cur,currency,string")
d=.TData("decimals,decimal")
c=TOT(c,TOT(j("cur"),TOT(Application.Contents("cur"),"GBP")))
If rxTest(c,z)Then
.Data("cur")=c
End If
If d="" Then
If rxTest(c,"^\w+$")Then
d=TOT(EXE("FWX_Cur_"&c&"_Decimals"),d)
End If
d=VOZ(TOT(d,2))
End If
.Data("_output")=d
End With:End Function
Public Function tag_cur():tag_cur=True:Call tag__String():With Me
Dim c,n,d,z:z="[A-Z]{3}$"
n=.TData("amount,number,string")
c=.TData("symbol,code,country,cur,currency")
d=.TData("decimals,decimal")
If rxTest(.Data("string"),"([0-9]+\.)?([0-9]+)?"&z)Then
c=rxSelect(.Data("string"),z)
n=rxReplace(.Data("string"),z,"")
End If
c=TOT(c,TOT(j("cur"),TOT(Application.Contents("cur"),"GBP")))
If rxTest(c,z)Then
.Data("cur")=c
If .BData("noSymbol")Then
.Data("symbol")=""
Else
If j("FWX_Cur_"&c)="" Then
j("FWX_Cur_"&c)=TOT(EXE("FWX_Cur_"&c),c)
End If
.Data("symbol")=j("FWX_Cur_"&c)
End If
End If
If d="" Then
If rxTest(c,"^\w+$")Then
If j("FWX_Cur_"&c&"_Decimals")="" Then
j("FWX_Cur_"&c&"_Decimals")=TOT(EXE("FWX_Cur_"&c&"_Decimals"),d)
End If
d=j("FWX_Cur_"&c&"_Decimals")
End If
d=VOZ(TOT(d,2))
End If
If n<>"" Then
If rxTest(n,"\d+")Then
.Data("len")=VOV(.NData("slice,len,decimals,accuracy"),d)
.Data("string")=n
.Data("number")=VOZ(n)
.Data("amount")=FormatNumber(n,.Data("len"))
.Data("display")=.Data("amount")
If IsTrue(.BData("abs"))Then
.Data("display")=FormatNumber(Abs(n),.Data("len"))
End If
End If
End If
.Data("_output")=.Data("symbol")&.Data("display")
End With:End Function
Public Function tag_cash():tag_cash=tag_currency:End Function
Public Function tag_money():tag_money=tag_currency:End Function
Public Function tag_currency():tag_currency=True:Call tag__Number():With Me
.Data("string")=FormatNumber(.Data("string"),2)
.Data("_output")=.Data("string")
End With:End Function
Public Function tag_num():With Me:Call tag__Number()
.Data("len")=VOZ(TOT(.TData("slice,len,decimals,accuracy"),2))
.Data("number")=FormatNumber(.Data("number"),.Data("len"))
.Data("_output")=.Data("number")
End With:End Function
Public Function tag_val():tag_val=tag_num:End Function
Public Function tag_round():tag_round=tag_num:End Function
Public Function tag_value():tag_value=tag_num:End Function
Public Function tag_number():tag_number=tag_num:End Function
Public Function tag_decimals():tag_decimals=tag_num:End Function
Public Function tag_fmtnumber():tag_fmtnumber=tag_num:End Function
Public Function tag_FormatNumber():tag_FormatNumber=tag_num:End Function
Public Function tag_percent():tag_percent=True:Call tag__Number():With Me
.Data("outof")=VOZ(TOT(.TData("outof,operand,100,max,from"),1))
.Data("decimals")=VOZ(TOT(.TData("decimals,decimal,accuracy"),1))
If .Data("string")>0 Then
.Data("percent")=(.Data("string")/.Data("outof"))*100
.Data("_output")=FormatNumber(.Data("percent"),.Data("decimals"))
End If
End With:End Function
Public Function tag_percentage():tag_percentage=tag_percent:End Function
Public Function tag_digits():tag_digits=True:Call tag__Number():With Me
If .Data("string")>0 Then
.Data("length")=.NData("length,len,digits")
.Data("_output")=Digits(.Data("string"),.Data("length"))
Else
If rxTest(.Data("_string"),"^[\d\,\.]+ [\d\,\.]+$")Then
Dim b:b=Split(.Data("_string")," ")
.Data("_output")=Digits(VOZ(b(0)),VOZ(b(1)))
Else
.Data("_output")=.Data("_string")
End If
End If
End With:End Function
Public Function tag_ccur():tag_ccur=True:Call tag_cur():With Me
.Data("_output")=ColourCodeNumber(.Data("amount"),.Data("_output"))
End With:End Function
Public Function tag_colourcur():tag_colourcur=tag_ccur:End Function
Public Function tag_cmoney():tag_cmoney=tag_ccur:End Function
Public Function tag_ccash():tag_ccash=tag_ccur:End Function
Public Function tag_net():tag_net=True:Call tag_cur():With Me
Dim RATE:RATE=VOV(.NData("rate,var,vat,tax,vartax"),xxxx_TAX_RATE)
.Data("amount")=.Data("amount")*(1 /(1+(RATE/100)))
Call tag_cur()
End With:End Function
Public Function tag_vat():tag_vat=True:Call tag_cur():With Me
Dim RATE:RATE=VOV(.NData("rate,var,vat,tax,vartax"),xxxx_TAX_RATE)
.Data("amount")=.Data("amount")*((RATE/100)/(1+(RATE/100)))
Call tag_cur()
End With:End Function
Public Function tag_gross():tag_gross=True:Call tag_cur():With Me
Dim RATE:RATE=VOV(.NData("rate,var,vat,tax,vartax"),xxxx_TAX_RATE)
.Data("amount")=.Data("amount")*(1+(RATE/100))
Call tag_cur()
End With:End Function
Public Function tag_phone():tag_phone=True:Call tag__String():With Me
.Data("_output")=DoPhone(.Data("string"))
End With:End Function
Public Function tag_tel():tag_tel=True:Call tag__String():With Me
Dim s:s=.TData("string,phone,primary,telephone,phone1,phone2,alternative")
Dim c:c=Country00(TOT(.TData("country,code"),FWX.Setting("country")))
Dim p:p=.TData("plus,p")
s=rxReplace(s,"^[0\+\s]+("&c&")?","")
s=rxReplace(s,"[^0-9]+","")
s=rxReplace(s,"^0","")
.Data("phone")=s
.Data("code")=c
.Data("plus")=p
.Data("_output")=p&c&s
End With:End Function
Public Function tag_email():tag_email=True:Call tag__String():With Me
.Data("_output")=DoEmail(.Data("string"))
End With:End Function
Public Function tag_pc():tag_pc=True:Call tag__String():With Me
.Data("_output")=DoPostcode(.Data("string"))
End With:End Function
Public Function tag_postcode():tag_postcode=True:Call tag__String():With Me
.Data("_output")=UCase(.Data("string"))
End With:End Function
Public Function tag_filename():tag_filename=True:Call tag__String():With Me
.Data("string")=rxReplace(.Data("string"),"(\&)","and")
.Data("string")=rxReplace(.Data("string"),"(^\W)|(\W$)","")
.Data("string")=rxReplace(.Data("string"),"\W+","-")
.Data("_output")=.Data("string")
End With:End Function
Public Function tag_thumb():tag_thumb=True:Call tag__String():With Me
.Data("i")=.NData("src,image,img,thumb,thumbnail,file,path,string")
If .Data("string")<>"" Then
.Data("w")=.NData("w,width,size,dimension")
.Data("h")=.NData("h,height,size,dimension")
If Left(.Data("i"),1)="/" Then
.Data("i")="/"&.Data("w")&"x"&.Data("w")&.Data("i")
Else
End If
End If
.Data("_output")=.Data("i")
End With:End Function
Public Function tag_URLDecode():tag_URLDecode=True:Call tag__String():With Me
.Data("_output")=URLDecode(.TData("string,rawInput"))
End With:End Function
Public Function tag_urlencode():tag_urlencode=True:Call tag__String():With Me
.Data("_output")=URLEncode(.TData("string,rawInput"))
End With:End Function
Public Function tag_decode():tag_decode=True:Call tag__String():With Me
.Data("_output")=Decode(.TData("string,rawInput"))
End With:End Function
Public Function tag_encode():tag_encode=True:Call tag__String():With Me
.Data("_output")=Encode(.TData("string,rawInput"))
End With:End Function
Public Function tag_htmlencode():tag_htmlencode=tag_encode:End Function
Public Function tag_htmldecode():tag_htmldecode=tag_decode:End Function
Public Function tag_text2html():tag_text2html=True:Call tag__String():With Me
.Data("_output")=Text2Html(.TData("string,rawInput"))
End With:End Function
Public Function tag_2html():tag_2html=tag_text2html:End Function
Public Function tag_pre():tag_pre=tag_text2html:End Function
Public Function tag_fixed():tag_fixed=tag_text2html:End Function
Public Function tag_slice():tag_slice=True:Call tag__String():With Me
Dim s,k,m
s=.TData("text,subject,string,a,b,c,d,e,f,g")
k=VOV(.NData("length,max,limit,len,l,slice,words,phrases,sentences"),50)
m=LCase(.TData("mode,method,type"))
If m="" Then m=IIf(.TData("phrases")<>"","phrases","")
If m="" Then m=IIf(.TData("letters")<>"","letters","")
If m="" Then m=IIf(.TData("words")<>"","words","")
If m="" Then m="w"
Select Case LCase(m)
Case "p","phrase","phrases":
s=SlicePhrases(s,k)
Case "w","word","words":
s=SliceWords(s,k)
Case Else:
s=Slice(s,k)
End Select
.Data("_output")=s
End With:End Function
Public Function tag_left():tag_left=True:Call tag__String():With Me
.Data("_output")=Left(.Data("string"),VOV(VOV(.Data("length"),.Data("max")),50))
End With:End Function
Public Function tag_right():tag_right=True:Call tag__String():With Me
.Data("_output")=Right(.Data("string"),VOV(VOV(.Data("length"),.Data("max")),50))
End With:End Function
Public Function tag_set():tag_set=tag_notempty:End Function
Public Function tag_ifset():tag_ifset=tag_notempty:End Function
Public Function tag_isset():tag_isset=tag_notempty:End Function
Public Function tag_notempty():tag_notempty=True:Call tag__iif():Call tag_tot():With Me
.Data("string")=.TData("string,cond,test")
.Data("cond")=CBool(.Data("string")<>"")
Call tag_iif()
End With:End Function
Public Function tag_isempty():tag_isempty=tag_empty:End Function
Public Function tag_empty():tag_empty=True:Call tag__iif():Call tag_tot():With Me
.Data("string")=.TData("string,cond,test")
.Data("cond")=CBool(.Data("string")="")
Call tag_iif()
End With:End Function
Public Function tag_nltz():tag_nltz=tag_gtz:End Function
Public Function tag_gtz():tag_gtz=True:Call tag__iif():Call tag_vov():With Me
.Data("string")=.NData("string,cond,test")
.Data("cond")=CBool(GTZ(.Data("string")))
Call tag_iif()
End With:End Function
Public Function tag_ngtz():tag_ngtz=tag_ltz:End Function
Public Function tag_ltz():tag_ltz=True:Call tag__iif():Call tag_vov():With Me
.Data("string")=.NData("string,cond,test")
.Data("cond")=CBool(Not GTZ(.Data("string")))
Call tag_iif()
End With:End Function
Public Function tag_notzero():tag_notzero=tag_nz:End Function
Public Function tag_nz():tag_nz=True:Call tag__iif():Call tag_vov():With Me
.Data("string")=.NData("string,cond,test")
.Data("cond")=CBool(VOZ(.Data("string"))<>0)
Call tag_iif()
End With:End Function
Public Function tag_izzero():tag_izzero=tag_iz:End Function
Public Function tag_iz():tag_iz=True:Call tag__iif():Call tag_vov():With Me
.Data("string")=.NData("string,cond,test")
.Data("cond")=CBool(VOZ(.Data("string"))=0)
Call tag_iif()
End With:End Function
Public Function tag_plural():tag_plural=True:Call tag__iif():Call tag_vov():With Me
.Data("string")=.NData("string,cond,test")
.Data("_output")=IIf(VOZ(.Data("string"))=1,"","s")
End With:End Function
Public Function tag__iif():With Me
If Not GTZ(Me.PropertyCount)Then
If rxTest(.Data("args"),"^[^;]*;[^;]*;[^;]*$")Then
Me.PropertyCount=3
Dim a
a=Split(.Data("args"),";")
.Data("cond")=a(0)
.Data("true")=a(1)
.Data("false")=a(2)
End If
End If
End With:End Function
Public Function tag_startswith():tag_startswith=True:Call tag__iif():Call tag_tot():With Me
.Data("test")=.TData("test,with,find,match,start,starts")
If .Data("test")="" Or .Data("string")="" Then
.Data("cond")="n"
ElseIf LCase(.Data("test"))=LCase(.Data("string"))Then
.Data("cond")="y"
Else
.Data("cond")=StartsWith(.Data("string"),.Data("test"))
If(Not IsTrue(.Data("cond")))And(IsTrue(.Data("acceptReverse"))Or IsEmpty(.Data("acceptReverse")))Then
.Data("cond")=StartsWith(.Data("test"),.Data("string"))
End If
End If
.Data("cond")=TOT(.Data("cond"),"n")
Call tag_iif()
End With:End Function
Public Function tag_endswith():tag_endswith=True:Call tag__iif():Call tag_tot():With Me
.Data("test")=.TData("test,with,find,match,end,ends")
If .Data("test")="" Or .Data("string")="" Then
.Data("cond")="n"
ElseIf LCase(.Data("test"))=LCase(.Data("string"))Then
.Data("cond")="y"
Else
.Data("cond")=EndsWith(.Data("string"),.Data("test"))
If(Not IsTrue(.Data("cond")))And(IsTrue(.Data("acceptReverse"))Or IsEmpty(.Data("acceptReverse")))Then
.Data("cond")=EndsWith(.Data("test"),.Data("string"))
End If
End If
.Data("cond")=TOT(.Data("cond"),"n")
Call tag_iif()
End With:End Function
Public Function tag_IsInList():tag_IsInList=tag_InList:End Function
Public Function tag_InList():tag_InList=True:Call tag__iif():Call tag_tot():With Me
.Data("string")=.TData("find,test,match,string")
.Data("inlist")=.TData("inlist,isinlist,InList,IsInList,list")
If .Data("inlist")<>"" And .Data("string")<>"" Then
.Data("cond")=IsInList(.Data("string"),.Data("inlist"))
End If
If .Data("y")="" And .Data("n")="" Then
.Data("y")="y"
End If
Call tag_iif()
End With:End Function
Public Function tag_checked():tag_checked=True:Call tag__iif():Call tag_tot():With Me
If .Data("or-dash,if-dash")="-" Then .Data("string")="y"
.Data("string")=.BData("string,cond,test,args")
.Data("inlist")=.TData("inlist,isinlist,InList,IsInList,list")
If .Data("inlist")<>"" Then
.Data("cond")=IsInList(.Data("string"),.Data("inlist"))
Else
.Data("cond")=IsTrue(.Data("string"))
End If
If .Data("y")="" Then
.Data("y")="checked='checked'"
End If
Call tag_iif()
End With:End Function
Public Function tag_selected():tag_selected=True:Call tag__iif():Call tag_tot():With Me
If .Data("cond")="" Then
If .Data("match")<>"" And(.Data("match")=.Data("current")Or .Data("match")=.Data("string"))Then
.Data("string")="y"
ElseIf .Data("match")<>"" And(.Data("match")=.Data("match-a")Or .Data("match")=.Data("match-b")Or .Data("match")=.Data("match-c"))Then
.Data("string")="y"
ElseIf .Data("current")<>"" And(.Data("current")=.TData("a,b,c,d,e"))Then
.Data("string")="y"
Else
If .Data("or-dash,if-dash")="-" Then
.Data("string")="y"
Else
.Data("string")=.BData("string,cond,test,args")
End If
End If
.Data("cond")=IsTrue(.Data("string"))
End If
If .Data("y")="" Then
.Data("y")="selected='selected'"
End If
Call tag_iif()
End With:End Function
Public Function tag_Disabled():tag_Disabled=True:Call tag__iif():Call tag_tot():With Me
If .Data("or-dash,if-dash")="-" Then .Data("string")="y"
.Data("string")=.BData("string,cond,test,args")
.Data("cond")=IsTrue(.Data("string"))
.Data("y")="disabled='disabled'"
Call tag_iif()
End With:End Function
Public Function tag_iif():tag_iif=True:With Me
Dim a,b,c,f
b=False
Call tag__iif()
If .Data("cond")<>"" Then
If rxTest(.Data("args"),"^[^;]*;[^;]*;[^;]*$")Then
c=.Data("cond")
b=tag_iif_ParseCondition(c)
End If
End If
If Not b Then
bHardCond=False
For Each f In Split("cond,if,condition,test,or,else",",")
c=""
If Not b Then
c=.Data(f)
bHardCond=bHardCond Or CBool(c<>"")
b=tag_iif_ParseCondition(c)
End If
Next
If Not(b Or bHardCond)Then
For Each f In Split("a,b,c,d,e,f",",")
c=""
If Not b Then
c=.ParsedData(f)
b=tag_iif_ParseCondition(c)
End If
Next
End If
End If
If b Then
ysss=Join(Split("y,true,yes",","),"-*,")&"-*"&","&"*"&Join(Split("true,yes,y",","),",*")
.Data("true")=.TData("true,yes,y,output,text,then")
.Data("bsss")=.TData(Replace(ysss,"*","before"))
.Data("asss")=.TData(Replace(ysss,"*","after"))
.Data("_output")=.Data("bsss")&.Data("true")&.Data("asss")
Else
nsss=Join(Split("n,false,no",","),"-*,")&"-*"&","&"*"&Join(Split("n,false,no",","),",*")
.Data("false")=.TData("false,not,no,n")
.Data("bsss")=.TData(Replace(nsss,"*","before"))
.Data("asss")=.TData(Replace(nsss,"*","after"))
.Data("_output")=.Data("bsss")&.Data("false")&.Data("asss")
End If
End With:End Function
Public Function tag_iif_ParseCondition(ByVal f)
tag_iif_ParseCondition=False
If f<>"" Then
If IsTrue(f)Then
tag_iif_ParseCondition=True
Exit Function
End If
If Left(f,1)="=" Then
tag_iif_ParseCondition=EXE2(Mid(f,2))
Exit Function
End If
Dim a,b,c,d,x,y:y=""""""
For Each d In Split("<>,<=,>=,<,>,=",",")
a="":x="":y=""
If c="" Then
a=Split(f,d)
If UBound(a)=1 Then
x=a(0)
y=a(1)
Else
If Left(f,Len(d))=d Then
x=Replace(Mid(f,Len(d)+1),"""","""""")
ElseIf Right(f,Len(d))=d Then
x=Replace(Mid(f,1,Len(f)-Len(d)),"""","""""")
End If
End If
If x<>"" Then
If d="=" Or d="<>" Then
x=Replace(x,"""","""""")
y=Replace(y,"""","""""")
c=""""&x&""""&d&""""&y&""""
Else
c=VOZ(x)&d&VOZ(y)&""
End If
End If
End If
Next
If c="" Then c=f
b=EXE2(c)
If IsTrue(b)And Err.Number=0 Then
tag_iif_ParseCondition=True
End If
End If
End Function
Public Function tag_random():tag_random=True:Call tag_tot():With Me
Dim s,p,x,a,i,how_many,d,be,ae,we
s=.TData("_output,string,paragraphs")
If s<>"" Then
d=TOT(.TData("divider,separator,split"),",")
be=.ParsedData("before-each")
ae=.ParsedData("after-each")
we=Split(.ParsedData("wrap-each"),"*")
i=1
how_many=.NData("n,qty,quantity,many,how-many,how_many")
If .TData("mode,type,config")="paragraph" Then
Do While i<=how_many And s<>""
p=RandomParagraph(s)
If p<>"" Then
s=Replace(s,p,"")
Else
p=s
i=how_many
End If
If p<>"" Then
p=be&p&ae
If UBound(we)=1 Then p=we(0)&p&we(1)
x=x&p
End If
i=i+1
Loop
Else
Do While i<=how_many And s<>""
a=Split(s,d)
Randomize()
If UBound(a)>0 Then
p=a(Round(Rnd()*UBound(a)))
s=Replace(s,p,"")
s=Replace(s,d&d,d)
Else
p=s
i=how_many
End If
If p<>"" Then
p=be&p&ae
If UBound(we)=1 Then p=we(0)&p&we(1)
x=x&p
End If
i=i+1
Loop
End If
End If
.Data("_output")=x
End With:End Function
Public Function tag_FirstParagraph():tag_FirstParagraph=True:Call tag_tot():With Me
.Data("fp")=.Data("_output")
If .Data("fp")<>"" Then
.Data("fp")=rxSelect(.Data("fp"),"^[\n\r\s\S]+?</p>")
.Data("fp")=rxReplace(.Data("fp"),"^.*<p\b","<p")
End If
.Data("_output")=.Data("fp")
End With:End Function
Public Function tag_like():tag_like=tag_instr:End Function
Public Function tag_find():tag_find=tag_instr:End Function
Public Function tag_instr():tag_instr=True:Call tag_tot():With Me
.Data("string")=.TData("string,cond,test")
.Data("find")=.TData("find,pattern,search,match,select")
.Data("cond")=CBool(InStr(1,.Data("string"),.Data("find"))>0)
Call tag_iif()
End With:End Function
Public Function tag_replace():tag_replace=True:Call tag_tot():With Me
.Data("find")=.TData("find,pattern,search,match")
.Data("rplc")=.TData("rplc,replace")
.Data("_output")=Replace(.Data("string"),.Data("find"),.Data("rplc"))
End With:End Function
Public Function tag_rreplace():tag_rreplace=tag_rxreplace:End Function
Public Function tag_rxreplace():tag_rxreplace=True:Call tag_tot():With Me
.Data("find")=.TData("find,pattern,search,match,select")
.Data("rplc")=.TData("rplc,replace,replacement")
.Data("_output")=rxReplace(.Data("string"),.Data("find"),.Data("rplc"))
End With:End Function
Public Function tag_rxtest():tag_rxtest=True:Call tag_tot():With Me
.Data("find")=.TData("find,pattern,search,match,select")
.Data("cond")=rxTest(.Data("string"),.Data("find"))
Call tag_iif()
End With:End Function
Public Function tag_rx():tag_rx=tag_regexp:End Function
Public Function tag_rfind():tag_rfind=tag_regexp:End Function
Public Function tag_rxfind():tag_rxfind=tag_regexp:End Function
Public Function tag_rselect():tag_rselect=tag_regexp:End Function
Public Function tag_rxselect():tag_rxselect=tag_regexp:End Function
Public Function tag_regexp():tag_regexp=True:Call tag_tot():With Me
.Data("find")=.TData("find,pattern,search,match,select")
.Data("rplc")=.TData("rplc,replace")
If .Data("rplc")<>"" Then
.Data("_output")=rxReplace(.Data("string"),.Data("find"),.Data("rplc"))
ElseIf .Data("true")<>"" Or .Data("false")<>"" Then
.Data("cond")=rxTest(.Data("string"),.Data("find"))
Call tag_iif()
Else
.Data("_output")=rxFind(.Data("string"),.Data("find"))
End If
End With:End Function
Public Function tag_each():tag_each=True:Call tag_tot():With Me
.Data("datasrc")=.TData("datasrc,source,object,collection,col")
If IsSet(.Data("datasrc"))Then
Else
tag_each=tag_split
Exit Function
End If
Call tag__DATA_Prepare_Template()
.Data("each")=.TData("each,template")
Dim out:Set out=New QuickString:out.Divider=.Data("join")
Dim src:Set src=OBJ(.Data("datasrc"))
Dim txt,i
For Each i In src
txt=Trim(src(i))
If txt<>"" Or .Data("AllowNulls")<>"" Then
tmp=.Data("each")
If tmp<>"" Then
.Data("name")=i:.Data("item")=i
.Data("value")=txt:.Data("i")=txt
If .Data("mark")<>"" Then tmp=Replace(tmp,.Data("mark"),txt)
txt=MergeTemplate(tmp,.Data)
End If
out.Add TOT(txt,"")
End If
Next
.Data("_output")=out
End With:End Function
Public Function tag_split():tag_split=True:Call tag_tot():With Me
.Data("split")=TOT(.TData("split,delimiter,divider,div"),",")
.Data("join")=.TData("join,divider")
Call tag__DATA_Prepare_Template()
.Data("each")=.TData("each,template")
.Data("mark")=.TData("mark,marker,tag,rplc,replace")
Dim out:Set out=New QuickString:out.Divider=.Data("join")
Dim src:src=Split(.Data("string"),.Data("split"))
Dim txt,i
Dim be:be=.TData("before-each,beforeEach,beforeeach")
Dim ae:ae=.TData("after-each,afterEach,aftereach")
Dim we:we=.TData("wrap-each,wrapEach,wrapeach")
Do While i<=UBound(src)
txt=Trim(src(i))
If txt<>"" Or .Data("AllowNulls")<>"" Then
tmp=.Data("each")
If tmp<>"" Then
.Data("n")=i+1:.Data("count")=.Data("n")
.Data("i")=txt:.Data("name")=txt:.Data("item")=txt
If .Data("mark")<>"" Then tmp=Replace(tmp,.Data("mark"),txt)
txt=MergeTemplate(tmp,.Data)
txt=Enclose(txt,be,ae)
If we<>"" Then
If InStr(1,we,"*")>0 Then
txt=Replace(we,"*",txt)
Else
we=""
End If
End If
End If
out.Add TOT(txt,"")
End If
i=i+1
Loop
.Data("_output")=out
End With:End Function
Public Function tag_for():tag_for=True:With Me
Dim o:Set o=New QuickString
.Data("i")=.NData("i,1,x,a,a,start,from,after")
.Data("j")=.NData("j,2,y,b,z,end,until,to,before")
If .Data("j")<.Data("i")Then .Data("j")=.Data("i")
Call tag__DATA_Prepare_Template()
Do While .Data("i")<=.Data("j")
o.Add MergeTemplate(.Data("template"),.Data)
.Data("i")=.Data("i")+1
Loop
.Data("_output")=o
End With:End Function
Public Function tag_progsplit():tag_progsplit=True:Call tag_tot():With Me
.Data("split")=TOT(.TData("split,delimiter,divider,div"),",")
.Data("pjoin")=TOT(.TData("progjoin,prog_join,pathdivider"),"_")
Dim out:Set out=New QuickString:out.Divider=.Data("split")
Dim src:src=Split(.Data("string"),.Data("split"))
Dim txt,i,pat
Do While i<=UBound(src)
txt=Trim(src(i))
If txt<>"" Or .Data("AllowNulls")<>"" Then
out.Add pat&txt
pat=pat&txt&.Data("pjoin")
End If
i=i+1
Loop
.Data("string")=out
Call tag_split()
End With:End Function
Public Function tag_tot():tag_tot=True:With Me
.Data("string")=Me.TData("string,str,text,txt,a,b,c,d,e")
.Data("_output")=.Data("string")
End With:End Function
Public Function tag_toe():tag_toe=tag_tot:End Function
Public Function tag_vov():tag_vov=True:With Me
.Data("string")=Me.NData("string,str,text,txt,a,b,c,d,e,alt")
If .Data("string")=0 Then
If Me.ClassName="voz" Then .Data("string")=VOZ(.ParsedData("args"))
End If
.Data("_output")=.Data("string")
End With:End Function
Public Function tag_voz():tag_voz=tag_vov:End Function
Public Function tag_fixed_field():With Me
Call tag__String()
.Data("L")=.NData("L,length,size")
If .Data("L")=0 Then
.Data("L")=(VOZ(.TData("stop,end")))-VOZ(.TData("start,begin"))+1
End If
If .Data("L")=0 Then
.Data("L")=1
End If
.Data("D")=.TData("a,b,c,d,e,f,g,h,D,DATA,data,string")
If .Data("removePrices")<>"" Then
.Data("D")=Decode(.Data("D"))
.Data("D")=Replace(.Data("D"),"&#xa3;","$")
.Data("D")=Replace(.Data("D"),vbCRLf," ")
.Data("D")=Replace(.Data("D"),vbCR," ")
.Data("D")=Replace(.Data("D"),vbLF," ")
.Data("D")=rxReplace(.Data("D"),"\s+"," ")
.Data("D")=rxReplace(.Data("D"),"\s*((for|\@) )?\$[\d\.]+[\s\r\n]*","")
.Data("D")=Replace(.Data("D"),"1x ","")
End If
Select Case .Data("D")
Case "@SPACE":.Data("D")=" "
Case "@CR":.Data("D")=vbCR
Case "@LF":.Data("D")=vbLF
Case "@INVOICE_ITEMS":.Data("D")=shop.Invoices.FieldByID("dbo.fncSHOP_Invoice_ItemCodes(t.invoice,'')",VOZ(.Data("invoice")))
End Select
.Data("X")=Left(.Data("D"),.Data("L"))
.Data("spaceReq")=VOZ(.Data("L")-Len(.Data("X")))
.Data("spaceChr")=TOT(.TData("space,spacechar,spaceChar,spaceChr"),"@SPACE")
If .Data("spaceReq")>0 Then
.Data("P")=String(.Data("spaceReq"),Replace(.Data("spaceChr"),"@SPACE"," "))
End If
If .Data("align")="right" Then
.Data("X")=.Data("P")&.Data("X")
Else
.Data("X")=.Data("X")&.Data("P")
End If
.Data("_output")=Replace(.Data("X")," ","@SPACE@")'&"             "&.Data("notes")&"."&vbCRLf
End With:End Function
Public Function tag_str():With Me:Call tag__String()
.Data("find")=.TData("find,pattern,search,match,select")
.Data("rplc")=.TData("rplc,replace")
If .Data("find")<>"" Then
If .Data("rplc")<>"" Then
.Data("string")=Replace(.Data("string"),.Data("find"),.Data("rplc"))
End If
End If
If GTZ(.Data("left"))Then
.Data("string")=Left(.Data("string"),VOZ(.Data("left")))
End If
If GTZ(.Data("right"))Then
.Data("string")=Left(.Data("string"),VOZ(.Data("left")))
End If
.Data("_output")=.Data("string")
End With:End Function
Public Function tag_text():tag_text=tag_str:End Function
Public Function tag_string():tag_string=tag_str:End Function
Public Function tag__String():With Me
.Data("string")=Me.TData("string,a,b,c,d,e")
If .Data("string")="" And .Data("args")<>"" Then
If rxTest(.Data("args"),"\=(\'|\"")")Then
Else
.Data("string")=.ParsedData("args")
End If
End If
.Data("LEFT")=.NData("left")
.Data("RIGHT")=.NData("right")
.Data("SLICE")=.NData("slice")
.Data("_string")=.Data("string")
End With:End Function
Public Function tag__Number():With Me
Call tag__String()
.Data("string")=Me.NData("number,a,b,c,d,e,string")
.Data("number")=VOZ(.Data("string"))
End With:End Function
Public Function tag__Date():With Me
.Data("str")=Me.TData("date,a,b,c,d,e,string")
.Data("min")=Me.TData("min")
.Data("max")=Me.TData("max")
.Data("str")=Me.DData("date,a,b,c,d,e,string")
.Data("min")=Me.DData("min")
.Data("max")=Me.DData("max")
.Data("str")=DateFromString(Me.DData("date,a,b,c,d,e,str"))
.Data("min")=DateFromString(Me.DData("min"))
.Data("max")=DateFromString(Me.DData("max"))
If IsDate(.Data("str"))Then
If IsDate(.Data("min"))Then
If CDate(.Data("str"))<CDate(.Data("min"))Then .Data("str")=.Data("min")
.Data("min-d")="y"
End If
If IsDate(.Data("max"))Then
If CDate(.Data("str"))>CDate(.Data("max"))Then .Data("str")=.Data("max")
.Data("max-d")="y"
End If
End If
End With:End Function
Function ParsedData(ByVal f)
ParsedData=ParseData(Me.Data(f))
End Function
Function ParseData(ByVal f)
ParseData=FWX_ParsedData(f)
End Function
Function TData(ByVal fs)
TData=""
Dim f,o:o=""
For Each f In Split(fs,",")
o=Me.ParsedData(f)
If o<>"" Then
TData=o
Exit Function
End If
Next
End Function
Function DData(ByVal fs)
DData=""
Dim f,so,o:o=""
For Each f In Split(fs,",")
so=Me.ParsedData(f)
o=so
If o<>"" Then
If IsDate(o)Then
DData=o
Exit Function
End If
If VOZ(o)<>0 Then
If Len(o)>10 Then
DData=DateAdd("s",VOZ(o),"1 Jan 1970 00:00")
Exit Function
ElseIf Abs(VOZ(o))<=365 Then
DData=DateAdd("d",VOZ(o),Date())
Exit Function
End If
End If
o=MakeDateFromString(o)
If o<>"" Then
DData=o
Exit Function
End If
End If
Next
End Function
Function NData(ByVal fs)
NData=0
Dim f,o:o=0
For Each f In Split(fs,",")
o=VOZ(Me.ParsedData(f))
If o<>0 Then
NData=o
Exit Function
End If
Next
End Function
Function BData(ByVal fs)
BData=False
Dim f,o:o=""
For Each f In Split(fs,",")
o=Me.ParsedData(f)
If IsTrue(o)Then
BData=True
Exit Function
End If
Next
End Function
Function JData(ByVal fs,ByVal d)
Dim f,o:o=""
Dim x:Set x=New QuickString:x.Divider=d
For Each f In Split(fs,",")
o=Me.ParsedData(f)
If o<>"" Then
x.Add o
End If
Next
JData=x
End Function
Function tab_album():tag_album=tag_picasa_album:End Function
Public Function tag_picasa_album():tag_picasa_album=True:With parser.oTag
Call tag__DATA()
.Data("picasa")="y"
DBG "tag_picasa_album"
.Data("width")=.NData("width,w,imgmax")
.Data("height")=.NData("height,h,imgmax")
.Data("imgmax")=.NData("imgmax,size,maxsize")
If .Data("imgmax")=0 Then
If .Data("width")>0 Then .Data("imgmax")=.Data("width")
If .Data("height")>0 And .Data("height")>.Data("width")Then .Data("imgmax")=.Data("height")
End If
If .Data("imgmax")>0 Then
.Data("imgmax-query")="?imgmax="&.Data("imgmax")
End If
Dim t
For Each t In Split("width,height,imgmax,imgmax-query",",")
.Data("template")=Replace(.Data("template"),"|"&t&"|",.Data(t))
Next
.Data("xml")=.TData("xml,url,link,album,href")
.Data("user")=.TData("user,username,account,userID")
If .Data("user")="" Then
.data("user")=site.CFG("google.username")
End If
If .Data("xml")="" Then
.Data("xml")="http://picasaweb.google.com/data/feed/base/user/"&.Data("user")&"?alt=rss&kind=photo&"
.Data("tags")=.TData("tags,key")
If .Data("tags")<>"" Then
Dim k:k=""
Dim q:Set q=New Quickstring:q.Divider="+OR+"
For Each k In Split(.Data("tags"),";")
k=Trim(k)
If k<>"" Then
q.Add """"&URLEncode(k)&""""
End If
Next
.Data("q")=q
.Data("xml")=.Data("xml")&"q="&.Data("q")&"&"
End If
End If
Call .tag_xml()
DBG "tag_picasa_album :: PROCESSED "&Len(.Data("_output"))&"b"
j("__picasa-album")=VOZ(j("__picasa-album"))+1
.Data("id")=TOT(.TData("id"),"picasa-album-"&j("__picasa-album"))
.Data("class")=TOT(.TData("class"),"picasa-album")
.Data("slide")=TOT(.TData("slide,slideshow"),"")
.Data("float")=Enclose(.TData("float,align"),"float:",";")
.Data("border")=Enclose(.TData("border,borders"),"border:",";")
.Data("background")=Enclose(.TData("background,bg"),"background:",";")
If .Data("slide")<>"" Then
.Data("slide")=" slide"
If .Data("width")>0 Then .Data("swidth")="width:"&.Data("width")&"px;"
If .Data("height")>0 Then .Data("sheight")="height:"&.Data("height")&"px;"
End If
.Data("_output")="<div id='"&.Data("id")&"' class='"&.Data("class")&.Data("slide")&"' style='"&.Data("float")&.Data("swidth")&.Data("sheight")&.Data("border")&.Data("background")&.Data("style")&"'"&IIf(IsTest," data-xml='"&.Data("xml")&"'","")&">"&.Data("_output")&"</div>"
DBG "tag_picasa_album :: FINALIZED "&Len(.Data("_output"))&"b"
End With:End Function
Public Function tag__Html():With Me
If .Data("url")="" Then .Data("url")=.Data("file")
If .Data("url")="" Then .Data("url")=.Data("source")
If .Data("url")="" Then .Data("url")=.Data("src")
If .Data("url")="" Then .Data("url")=.Data("args")
End With:End Function
Public Function tag_bootstrap():Call tag__String():With Me
Dim v,p,f,t,g,x
For Each v In Split("string,file,path,virtual,tmpl,template,a,b,c,d,e",",")
p=.Data(v)
If p<>"" And t="" Then
f=p
x=rxTest(f,"\.html?$")
If StartsWith(p,"/")Then
f=f&IIf(x,"/fyneworks.html","")
t=ReadFile(f)
If t<>"" Then
g=p
t=parser.PreParse(t)
End If
Else
t=GetTemplate(f)
If t<>"" Then
g=p
End If
End If
End If
Next
If t<>"" And g<>"" Then
t=rxReplace(t,"(href|src)=(""|')(\w+)\/","$1=$2"&g&"$3/")
End If
.Data("_output")=t
End With:End Function
Public Function tag__Skype():With Me
If .Data("id")="" Then .Data("id")=.Data("user")
If .Data("id")="" Then .Data("id")=.Data("email")
End With:End Function
Public Function tag_skype():tag_skype=True:Call tag__Skype():With Me
.Data("status")=tag__skype_status()
If .Data("status")="" Then
.Data("status")="Offline"
End If
If .Data("action")<>"" Then
.Data("__")=Eval("tag_skype_"&.Data("action")&"")
End If
If .Data("_output")="" Then
.Data("_output")=.Data("status")
End If
End With:End Function
Public Function tag_skype_if():tag_skype_if=True:Call tag__Skype():With Me
.Data("s")=LCase(RemoveSymbols(.Data("status")))
.Data("_output")=.Data(.Data("s")&"")
.Data("_output")=TOT(.Data("_output"),.Data("else"))
End With:End Function
Public Function tag_skype_status():tag_skype_status=True:Call tag__Skype():With Me
.Data("_output")=tag__skype_status()
End With:End Function
Public Function tag__skype_status():With Me
Err.Clear():On Error Resume Next
If j("SKYPE"&.Data("id")&"")="" Then
Dim objXMLHTTP:Set objXMLHTTP=Server.CreateObject("Microsoft.XMLHTTP")
objXMLHTTP.Open "GET","http://mystatus.skype.com/"&.Data("id")&".txt",False
objXMLHTTP.SetRequestHeader "Content-type","text/plain"
objXMLHTTP.Send()
j("SKYPE"&.Data("id")&"")=objXMLHTTP.ResponseText
Set objXMLHTTP=Nothing
End If
Err.Clear()
tag__skype_status=j("SKYPE"&.Data("id")&"")
End With:End Function
Public Function tag_youtube():tag_youtube=True:With Me
.Data("_output")="Hello YouTube"
End With:End Function
Public Function tag_ysm():tag_ysm=tag_YahooPPC:End Function
Public Function tag_YahooPPC():tag_YahooPPC=True
Dim o:Set o=New QuickString
Call tag__String()
Me.Data("account")=TOE(Me.TData("account,id,code,analytics,string"))
If Me.Data("account")="" Then Exit Function
Me.Data("amount")=Me.NData("amount,value,total")
Me.Data("currency")=TOT(Me.TData("currency"),"GBP")
Me.Data("invoice")=Me.NData("invoice,inv,sale,ecomm,ecommerce")
If Me.Data("invoice")<>0 Then
Set rInv=shop.AllInvoices.View.Find("t.invoice="&Me.Data("invoice")&" AND t.amount>0")
If Exists(rInv)Then
j("amount")=VOV(j("amount"),rInv("amount"))
End If
KillRS rInv
End If
o.Add "<scr"&"ipt type='text/javascript'>"
o.Add "window.ysm_customData = new Object();"
o.Add "window.ysm_customData.conversion = ""transId="&Me.Data("invoice")&",currency="&Me.Data("currency")&",amount="&Me.Data("amount")&""";"
o.Add "var ysm_accountid  = """&Me.Data("account")&""";"
o.Add "document.write(""<SCR"" + ""IPT language='Javascript' type='text/javascript' "" "
o.Add "+ ""SRC=//"" + ""srv1.wa.marketingsolutions.yahoo.com"" + ""/script/ScriptServlet"" + ""?aid="" + ysm_accountid "
o.Add "+ ""></SCR"" + ""IPT>"");"
o.Add "</scr"&"ipt>"
Me.Data("_output")=o
End Function
Public Function tag_dist():tag_dist=tag_distance:End Function
Public Function tag_distance():tag_distance=True:Call tag__String():With Me
.Data("from")=.TData("from,start,go,string")
.Data("to")=.TData("to,dest,destination,finish,end,place,target")
.Data("_output")=GetDistance(.Data("from"),.Data("to"))
.Data("a")=.Data("_output")
Call tag_Num()
End With:End Function
Public Function tag_GoogleMapsAPI():GoogleMapsAPI=True
Me.Data("_output")=GoogleMapsAPI()
End Function
Public Function tag_GoogleMapsKey():tag_GoogleMapsKey=True
Me.Data("_output")=GoogleMapsKey()
End Function
Public Function tag_map():tag_map=tag_GoogleMap:End Function
Public Function tag_gmap():tag_gmap=tag_GoogleMap:End Function
Public Function tag_GoogleMap():tag_GoogleMap=True
If Me.Data("lng")="" Or Me.Data("lat")="" Then
Me.Data("geo")=Me.TData("geo,where,pc,postcode,coords,coordinates,coord,position")
If Me.Data("geo")="" Then
Me.Data("_output")="<span style='color:red;'>missing postcode for Google map</span>"
Exit Function
End If
Me.Data("coords")=rxSelect(Me.Data("geo"),"[\d\.]+,[\d\.]+")
If Me.Data("coords")="" Then
Me.Data("coords")=GoogleCoords(Me.Data("geo"))
End If
If Me.Data("coords")="" Then
Me.Data("_output")="<span style='color:red;'>Coordinates not found for '"&Me.Data("geo")&"'</span>"
End If
Dim acrd:acrd=Split(Me.Data("coords"),",")
If UBound(acrd)=1 Then
Me.Data("lat")=acrd(0)
Me.Data("lng")=acrd(1)
End If
End If
If(Not Me.Data("key")<>"")Then Me.Data("key")=GoogleMapsKey()
If IsEmpty(Me.Data("zoom"))Then Me.Data("zoom")=13
If IsEmpty(Me.Data("width"))Then Me.Data("width")=500
If IsEmpty(Me.Data("height"))Then Me.Data("height")=400
If IsEmpty(Me.Data("mlabel"))Then Me.Data("mlabel")=Me.Data("label")
If rxTest(Me.Data("width"),"^[\d\.]+$")Then Me.Data("width")=Me.Data("width")&"px"
If rxTest(Me.Data("height"),"^[\d\.]+$")Then Me.Data("height")=Me.Data("height")&"px"
If Me.Data("dest")="" Then Me.Data("dest")=Me.Data("address")
If Me.Data("dest")="" Then Me.Data("dest")=Me.Data("postcode")
If Me.Data("dest")="" Then Me.Data("dest")=Me.Data("pc")
Me.Data("directions")=TOT(Me.Data("Directions"),Me.Data("directions"))
If Me.Data("directions")<>"" Then
Me.Data("mlabel")=Me.Data("mlabel")&"<div class='map-directions' style='font-size:9.0pt; clear:both;'><form action='http://maps.google.co.uk/maps' method='get' target='_blank'><div><input type='text' name='saddr' id='saddr' style='width:110px;' onfocus=""if(this.value==this.title){ this.value=''; }"" onblur=""if(this.value==''){ this.value=this.title; }"" onkeypress=""if(window.event.keyCode==13){ this.form.submit(); return false; }"" value='address or postcode' title='address or postcode'/></div><div><input type='submit' value='Get Directions'/></div><input type='hidden' name='daddr' value='"&Me.Data("dest")&"' /><input type='hidden' name='hl' value='en' /></div></form>	</div>	"
End If
If Me.Data("tfl")="" Then Me.Data("tfl")=Me.Data("JourneyPlanner")
If Me.Data("tfl")<>"" Then
Me.Data("mlabel")=Me.Data("mlabel")&"<div style='padding:3px 3px 3px 15px; clear:both;'><a href=' http://journeyplanner.tfl.gov.uk/user/XSLT_TRIP_REQUEST2?language=en&execInst=&sessionID=0&ptOptionsActive=-1&place_destination=London&name_destination="&Me.Data("dest")&"&type_destination=locator' target='_blank'><img hspace='0' vspace='0' src='/@/journey-planner.gif' title='Get London Transport directions with Journey Planner' /></a></div>"
End If
If Me.Data("mlat")="" Then Me.Data("mlat")=Me.Data("lat")
If Me.Data("mlng")="" Then Me.Data("mlng")=Me.Data("lng")
If Not GTZ(j("GoogleMap-count"))Then j("GoogleMap-count")=0
j("GoogleMap-count")=j("GoogleMap-count")+1
If Not(Me.Data("id")<>"")Then Me.Data("id")="GoogleMap_"&j("qid")&"_"&VOZ(Query.FQ("loader"))&"_"&j("GoogleMap-count")&""
Dim o:Set o=New QuickString
With o
.Add "<div class=""GoogleMap"" id="""&Me.Data("id")&""" style=""width: "&Me.Data("width")&"; height: "&Me.Data("height")&";"">"
.Add "<span style='display:block;padding:30px;text-align:center;'><img src='/@/wait.gif' alt='' height='16' width='16'/> Loading map...</span>"
.Add "</div>"
.Add "<script type=""text/javascript"">"
.Add "if(window.jQuery)(function($){$(function(){"
.Add "fwx.maps.load(function(){"
.Add "window['"&Me.Data("id")&"']=(function(id){"
If Me.Data("mlat")<>"" And Me.Data("mlng")<>"" Then
.Add "var center = new google.maps.LatLng("&VOZ(Me.Data("lat"))&", "&VOZ(Me.Data("lng"))&");"
End If
.Add "var map = new google.maps.Map(document.getElementById(id),{"
If Me.Data("mlat")<>"" And Me.Data("mlng")<>"" Then
.Add "center: center||null,"
End If
.Add "zoom: "&VOV(Me.Data("zoom"),13)&","
If Me.Data("controls")<>"" Then Me.Data("noui")=IIf(Me.Data("controls"),"n","y")
Me.Data("disableDefaultUI")=Me.TData("disableDefaultUI,disableUI,noUI,noui")
For Each setting In Split("disableDefaultUI,panControl,zoomControl,mapTypeControl,scaleControl,streetViewControl,overviewMapControl",",")
If Me.Data(setting)<>"" Then
.Add setting&": "&IIf(Me.Data(setting),"true","false")&","
End If
Next
.Add "mapTypeId: google.maps.MapTypeId."&TOT(Me.Data("type"),"ROADMAP")&""
.Add "});"
If Me.Data("mlat")<>"" And Me.Data("mlng")<>"" Then
.Add "var marker = new google.maps.Marker({"
.Add "map:map, position:center"
.Add "});"
If Me.Data("mlabel")<>"" Then
.Add "var infowindow = new google.maps.InfoWindow({"
.Add "content:unescape("""&Escape(Me.Data("mlabel"))&""")"
.Add "});"
.Add "google.maps.event.addListener(marker, 'click', function(){"
.Add "infowindow.open(map,marker);"
.Add "});"
End If
End If
.Add "return map;"
.Add "})('"&Me.Data("id")&"');"
.Add "});"
.Add "});})(jQuery);</script>"&vbCRLf
End With
Me.Data("_output")=o
End Function
Public Function tag_conversions():tag_conversions=True
Dim o
o=""
Me.Data("conversion")=Me.TData("action,label,goal,conversion")
If Me.Data("conversion")<>"" Then
Call tag_analytics
o=o&Me.Data("analytics")
Call tag_goals
o=o&Me.Data("goals")
End If
Me.Data("_output")=o
End Function
Public Function tag_goals():tag_goals=True
Call tag__String()
Dim o:Set o=New QuickString
o.Divider=vbCRLf
Dim c
For Each c In Split(Me.TData("string,label,action,conversions,conversion"),",")
c=TOE(c)
If c<>"" Then
If Me.Data("ga")<>"" Then o.Add "<script>/*ga*/$(function(){if(window['_gaq'])_gaq.push(['_trackPageview', '/do/google/adwords/conversion/"&c&"/']);});</script>"
If Me.Data("ua")<>"" Then o.Add "<script>/*ua*/$(function(){if(window['ga'])ga('send', 'pageview', '/do/google/adwords/conversion/"&c&"/');});</script>"
If IsTest Then
o.Add "<div class='Clear TCenter No B Small P20'>GOAL "&c&""
If Me.Data("ga")<>"" Then o.Add " <a href='/do/google/adwords/conversion/"&c&"' target='_blank'>GA(classic)</a>"
If Me.Data("ua")<>"" Then o.Add " <a href='/do/google/adwords/conversion/"&c&"' target='_blank'>UA(universal)</a>"
o.Add "</div>"
End If
Me.Data("label")=c
tag_conversion()
End If
Next
Me.Data("goals")=Me.Data("goals")&vbCRLf&o
Me.Data("goals")=Me.Data("goals")&vbCRLf&Me.Data("conversions")
Me.Data("_output")=o
End Function
Public Function tag_conversion():tag_conversion=True
Dim o:Set o=New QuickString
Me.Data("account")=TOE(Me.TData("adwords,account,id,code,conversion,adwords"))
If Me.Data("account")="" Then
If IsTest Then
o.Add "<div class='Clear TCenter No B Small'>Google Adwords CONVERSION - is OFF</div>"
End If
o.Add "<script>/* google adwords conversion ID missing */</script>"
Else
Me.Data("value")=Me.NData("value,amount,total")
Call tag__String()
Me.Data("label")=Me.TData("label,lbl,title,a,b,c,d,e,f,string")
If Me.Data("label")="" Then Me.Data("label")="enquiry"
Me.Data("action")=Me.TData("action,type,action,id")
If Me.Data("action")="" And Me.Data("label")<>"" Then
Me.Data("action")=site.CFG("google.adwords."&Me.Data("label"))
End If
If Me.Data("action")="" Then
If Me.Data("value")>0 Then
If IsTrue(Me.Data("do-not-default-to-purchase"))Then
ElseIf IsTrue(Me.Data("is-purchase"))Then
Me.Data("action")="purchase"
Else
Me.Data("action")=Me.Data("default-action")
End If
End If
End If
With o
If Me.Data("action")<>"" Then
If IsTest Then
.Add "<div class='Clear TCenter No B Small'>"
.Add "Google Adwords CONVERSION - "
.Add "account:"&Me.Data("account")&", "
.Add "action:"&Me.Data("action")&", "
.Add "value:"&Me.Data("value")&""
.Add "</div>"
End If
.Add "<script>"&vbCRLf
.Add "var google_conversion_id = """&Me.Data("account")&""";"&vbCRLf
.Add "var google_conversion_language = ""en_GB"";"&vbCRLf
.Add "var google_conversion_format = ""1"";"&vbCRLf
.Add "var google_conversion_color = ""ffffff"";"&vbCRLf
.Add "var google_conversion_value = "&VOZ(Me.Data("value"))&";"&vbCRLf
.Add "var google_conversion_label = """&Me.Data("action")&""";"&vbCRLf
.Add "</script>"&vbCRLf
.Add "<script type=""text/javascript"" src="""&MyProtocol&"://www.googleadservices.com/pagead/conversion.js"">"&vbCRLf
.Add "</script>"&vbCRLf
.Add "<noscript>"&vbCRLf
.Add "<img height=""1"" width=""1"" border=""0"" "
.Add "src="""&MyProtocol&"://www.googleadservices.com/pagead/conversion/"&Me.Data("account")&"/?"
.Add "value="&Me.Data("value")&"&"
.Add "label="&Me.Data("action")&"&amp;script=0""/>"&vbCRLf
.Add "</noscript>"&vbCRLf
.Add vbCRLf
Else
If IsTest Then
.Add "<div class='Clear TCenter No B Small'>"
.Add "Google Adwords NO CONVERSION at point of "&Me.Data("label")&""
.Add "</div>"
End If
.Add "<script>"&vbCRLf
.Add "/* no conversion to record here - "&vbCRLf&PrintColVars(Me.Data,"ga,ua,action,label,amount,is-purchase")&" */"
.Add "</script>"&vbCRLf
End If
End With
End If
Me.Data("conversions")=Me.Data("conversions")&o
Me.Data("_output")=o
End Function
Public Function tag_analytics_UNIVERSAL()
tag_analytics_UNIVERSAL=tag_analytics_UA
End Function
Public Function tag_analytics_UA():tag_analytics_UA=True
Call tag__String()
Me.Data("ua-account")=TOE(Me.TData("ua,universal,ua-account,id,code,analytics,string"))
If Me.Data("ua-account")="" Then Exit Function
Dim o:Set o=New QuickString
o.Divider=vbCRLf
o.Add "<scr"&"ipt>"
If j("analytics-ua-initialized")="" Then
j("analytics-ua-initialized")="y"
o.Add "(function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){"
o.Add "(i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),"
o.Add "m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)"
o.Add "})(window,document,'script','//www.google-analytics.com/analytics.js','ga');"
o.Add ""
o.Add "ga('set', 'forceSSL', true);"
End If
o.Add "ga('create', '"&Me.Data("ua-account")&"', {'cookieDomain': '"&IIf(IsTest,"none",MyDomain)&"'});"
o.Add "ga('require', 'displayfeatures');"
If GTZ(Me.Data("uid"))Then o.Add "ga('set', '&uid', '"&Me.Data("uid")&"');"
If Me.Data("conversion")="" Then o.Add "ga('send', 'pageview');"
Me.Data("invoice")=Me.NData("invoice,inv,sale,ecomm,ecommerce,id")
If Me.Data("invoice")<>0 And IsTrue(Me.Data("analytics-transaction"))Then
Set rInv=shop.AllInvoices.View.FindFields(_
"t.invoice,t.merchantName,t.amount,t.jobs_total,t.tax",_
"t.invoice="&Me.Data("invoice")&" AND t.active=1 AND t.amount>0" _
)
If Exists(rInv)Then
j("amount")=VOV(j("amount"),rInv("amount"))
o.Add "ga('require', 'ecommerce', 'ecommerce.js');"
o.Add "ga('set', 'currencyCode', 'GBP');"
o.Add "ga('ecommerce:addTransaction', {'id': '"&rInv("invoice")&"','affiliation': '"&EncodeJSON(rInv("merchantName"))&"','revenue': '"&FormatNumber(rInv("amount"),2)&"','shipping': '"&FormatNumber(rInv("jobs_total"),2)&"','tax': '"&FormatNumber(rInv("tax"),2)&"'});"
Set rItm=shop.AllInvoices.Items.View.FindFields(_
"t.invoice,t.name,t.code,t.fixprice,t.product,t.item,f.folderName,f.pageName,t.quantity",_
"t.invoice="&Me.Data("invoice")&" AND t.active=1" _
)
Do While Exists(rItm)
o.Add "ga('ecommerce:addItem', {'id': '"&rItm("invoice")&"','name': '"&EncodeJSON(rItm("name"))&"','sku': '"&TOT(rItm("code"),VOV(rItm("product"),rItm("item")))&"','category': '"&EncodeJSON(TOT(rItm("folderName"),rItm("pageName")))&"','price': '"&FormatNumber(rItm("fixprice"),2)&"','quantity': '"&rItm("quantity")&"'});"
rItm.MoveNext()
Loop
KillRS rItm
Set rItm=shop.AllInvoices.Jobs.View.FindFields(_
"t.invoice,t.name,t.code,t.fixprice,t.[option],t.job,f.folderName,s.name serviceName","t.invoice="&Me.Data("invoice")&" AND t.active=1")
Do While Exists(rItm)
o.Add "ga('ecommerce:addItem', {'id': '"&rItm("invoice")&"','name': '"&EncodeJSON(rItm("name"))&"','sku': '"&TOT(rItm("code"),VOV(rItm("option"),rItm("job")))&"','category': '"&EncodeJSON(TOT(rItm("folderName"),rItm("serviceName")))&"','price': '"&FormatNumber(rItm("fixprice"),2)&"','quantity': '1'});"
rItm.MoveNext()
Loop
KillRS rItm
o.Add ""
o.Add "ga('ecommerce:send');"
o.Add "ga('ecommerce:clear');"
End If
KillRS rInv
End If
o.Add "</scr"&"ipt>"
Me.Data("analytics-ua")=o
Me.Data("_output")=o
End Function
Public Function tag_analytics_CLASSIC()
tag_analytics_CLASSIC=tag_analytics_GA
End Function
Public Function tag_analytics_GA():tag_analytics_GA=True
Call tag__String()
Me.Data("ga-account")=TOE(Me.TData("ga,classic,ga-account"))
If Me.Data("ga-account")="" Then Exit Function
Dim o:Set o=New QuickString
o.Divider=vbCRLf
o.Add "<scr"&"ipt>"
o.Add "var _gaq = _gaq || [];"
o.Add "_gaq.push(['_setAccount', '"&Me.Data("ga-account")&"']);"
o.Add "_gaq.push(['_setDomainName', '"&MyDomain&"']);"
o.Add "_gaq.push(['_setAllowLinker', true]);"
If Me.Data("conversion")="" Then o.Add "_gaq.push(['_trackPageview']);"
Me.Data("invoice")=Me.NData("invoice,inv,sale,ecomm,ecommerce")
If Me.Data("invoice")<>0 Then
Set rInv=shop.AllInvoices.View.FindFields(_
"t.invoice,t.merchantName,a.city,a.county,a.country,t.jobs_total,t.amount,t.tax",_
"t.invoice="&Me.Data("invoice")&" AND t.active=1 AND t.amount>0" _
)
If Exists(rInv)Then
j("amount")=VOV(j("amount"),rInv("amount"))
o.Add "_gaq.push(['_addTrans','"&rInv("invoice")&"','"&EncodeJSON(TOT(rInv("merchantName"),shop.CFG("name")))&"','"&FormatNumber(rInv("amount"),2)&"','"&FormatNumber(rInv("tax"),2)&"','"&rInv("jobs_total")&"','"&EncodeJSON(rInv("city"))&"','"&rInv("county")&"','"&rInv("country")&"']);"
Set rItm=shop.AllInvoices.Items.View.FindFields(_
"t.invoice,t.name,t.code,t.fixprice,t.product,t.item,f.folderName,f.pageName,t.quantity",_
"t.invoice="&Me.Data("invoice")&" AND t.active=1" _
)
Do While Exists(rItm)
o.Add "_gaq.push(['_addItem','"&rItm("invoice")&"','"&TOT(rItm("code"),VOV(rItm("product"),rItm("item")))&"','"&EncodeJSON(rItm("name"))&"','"&EncodeJSON(rItm("folderName"))&"','"&FormatNumber(rItm("fixprice"),2)&"','"&rItm("quantity")&"']);"
rItm.MoveNext()
Loop
KillRS rItm
Set rItm=shop.AllInvoices.Jobs.View.FindFields(_
"t.invoice,t.name,t.code,t.fixprice,t.[option],t.job,f.folderName,s.name serviceName","t.invoice="&Me.Data("invoice")&" AND t.active=1")
Do While Exists(rItm)
o.Add "_gaq.push(['_addItem','"&rItm("invoice")&"','"&TOT(rItm("code"),VOV(rItm("option"),rItm("job")))&"','"&EncodeJSON(rItm("name"))&"','"&EncodeJSON(TOT(rItm("folderName"),rItm("serviceName")))&"','"&FormatNumber(rItm("fixprice"),2)&"','"& 1&"']);"
rItm.MoveNext()
Loop
KillRS rItm
o.Add "_gaq.push(['_trackTrans']);"
End If
KillRS rInv
End If
If j("analytics-ga-initialized")="" Then
j("analytics-ga-initialized")="y"
o.Add "(function(){"
o.Add "var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;"
o.Add "ga.src = ('https:' == document.location.protocol ? 'https://' : 'http://') + 'stats.g.doubleclick.net/dc.js';"
o.Add "var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);"
o.Add "})();"
End If
o.Add "</scr"&"ipt>"
Me.Data("analytics-ga")=o
Me.Data("_output")=o
End Function
Public Function tag_analytics():tag_analytics=True
Dim o
o=""
Call tag_analytics_GA
o=o&Me.Data("analytics-ga")
Call tag_analytics_UA
o=o&Me.Data("analytics-ua")
Me.data("analytics")=o
Me.Data("_output")=o
End Function
Public Function tag_spelling():tag_spelling=tag_GoogleSpelling:End Function
Public Function tag_gspelling():tag_gspelling=tag_GoogleSpelling:End Function
Public Function tag_GoogleSpelling():tag_GoogleSpelling=True
Call tag__String()
If Me.Data("string")<>"" Then
Me.Data("_output")=GoogleSpelling(Me.Data("string"))
End If
End Function
Public Function tag_doc():tag_doc=tag_gdocs:End Function
Public Function tag_gdoc():tag_gdoc=tag_gdocs:End Function
Public Function tag_gdocs():tag_gdocs=True
Dim o:Set o=New QuickString
Call tag__String()
Me.Data("doc")=TOE(Me.TData("doc,document,docid,id,string"))
If Me.Data("doc")="" Then
Me.Data("_output")="Missing Google Document ID"
Exit Function
End If
o=GetResponse(Me.Data("doc"))
o=Replace(o,"ifq","")
o=rxSelect(o,"<form[^>]*?>[\s\S]*?</form>")
o="<div id='gdocs'>"&o&"</div>"
Me.Data("_output")=o
End Function
Public Function tag_ecomm_prodid():With Me:Call tag_tot()
.Data("x")=.TData("prodid,ecomm_prodid,string")
If .Data("x")<>"" Then
.Data("x")=rxReplace(.Data("x"),"([\w\-\_\d]+)","'$1'")
.Data("x")=rxReplace(.Data("x"),"'+","'")
.Data("x")=rxReplace(.Data("x"),"^'|'$","'")
.Data("_output")=", ecomm_prodid: ["&.Data("x")&"]"
End If
End With:End Function
Public Function tag_ecomm_pvalue():With Me:Call tag_tot()
.Data("x")=.TData("pvalue,ecomm_pvalue,string")
If .Data("x")<>"" Then
.Data("_output")=", ecomm_pvalue: ["&.Data("x")&"]"
End If
End With:End Function
Public Function tag_ecomm_totalvalue():With Me:Call tag_tot()
.Data("x")=.TData("totalvalue,ecomm_totalvalue,string")
If .Data("x")<>"" Then
.Data("_output")=", ecomm_totalvalue: '"&.Data("x")&"'"
End If
End With:End Function
Public Function tag_ecomm_pcat():With Me:Call tag_tot()
.Data("x")=.TData("pcat,ecomm_pcat,string")
If .Data("x")<>"" Then
.Data("x")=rxReplace(.Data("x"),"[\'\""\n\r\t]+","")
.Data("x")=rxReplace(.Data("x"),"([^,]+)","'$1'")
.Data("x")=rxReplace(.Data("x"),"'\s+|\s+'","'")
.Data("x")=rxReplace(.Data("x"),"'+","'")
.Data("x")=rxReplace(.Data("x"),"^'|'$","'")
.Data("_output")=", ecomm_pcat: ["&.Data("x")&"]"
End If
End With:End Function
Public Function tag_encrypt():tag_encrypt=True:Call tag__String():With Me
.Data("_output")=Encrypt(.Data("string"))
End With:End Function
Public Function tag_enc():tag_enc=tag_encrypt:End Function
Public Function tag_decrypt():tag_decrypt=True:Call tag__String():With Me
.Data("_output")=Encrypt(.Data("string"))
End With:End Function
Public Function tag_dec():tag_dec=tag_decrypt:End Function
Public Function tag_md5():tag_md5=True:Call tag__String():With Me
.Data("_output")=MD5(.Data("string"))
End With:End Function
Public Function tag_ANNumber():tag_ANNumber=True:Call tag__String():With Me
.Data("_output")=ANNumber(.Data("string"))
End With:End Function
Public Function tag_ANString():tag_ANString=True:Call tag__String():With Me
.Data("_output")=ANString(.Data("string"))
End With:End Function
Public Function tag_sha():tag_sha=True:Call tag__String():With Me
.Data("_output")=SHA256(.Data("string"))
End With:End Function
Public Function tag_gravatar():tag_gravatar=True:Call tag__String():With Me
.Data("_output")="//www.gravatar.com/avatar/"&LCase(MD5(.Data("string")))
End With:End Function
Public Function tag_SeasonedHash():tag_SeasonedHash=True:With Me
.Data("_output")=SeasonedHash(.Data("password"),.Data("salt"))
End With:End Function
Public Function tag_requery():tag_requery=tag_ReAssignQuery:End Function
Public Function tag_ReAssignQuery():tag_ReAssignQuery=True:Call tag__String():With Me
.Data("_output")=ReAssignQuery(.Data("string"))
End With:End Function
Public Function tag_resubmitRequest():tag_resubmitRequest=tag_resubmit:End Function
Public Function tag_resubmitVars():tag_resubmitVars=tag_resubmit:End Function
Public Function tag_ReSubmitAll():tag_ReSubmitAll=tag_resubmit:End Function
Public Function tag_ReSubmitQry():tag_ReSubmitQry=tag_resubmit:End Function
Public Function tag_resubmit():tag_resubmit=True:Call tag__String():With Me
.Data("_output")=ReSubmit(.Data("string"))
End With:End Function
Public Function tag_ReSubmitQuery():tag_ReSubmitQuery=True:Call tag__String():With Me
.Data("_output")=ReSubmitQuery(.Data("string"))
End With:End Function
Public Function tag_ReAssignRequest():tag_ReAssignRequest=tag_ReAssign:End Function
Public Function tag_ReAssignForm():tag_ReAssignForm=tag_ReAssign:End Function
Public Function tag_ReAssignVars():tag_ReAssignVars=tag_ReAssign:End Function
Public Function tag_ReAssignAll():tag_ReAssignAll=tag_ReAssign:End Function
Public Function tag_repost():tag_repost=tag_ReAssign:End Function
Public Function tag_ReAssign():tag_ReAssign=True:Call tag__String():With Me
.Data("_output")=ReAssign(.Data("string"))
End With:End Function
Public Function tag_ReloadForm():tag_ReloadForm=tag_ReSubmitFields:End Function
Public Function tag_ReloadFields():tag_ReloadFields=tag_ReSubmitFields:End Function
Public Function tag_resubmitForm():tag_resubmitForm=tag_ReSubmitFields:End Function
Public Function tag_ReSubmitFields():tag_ReSubmitFields=True:Call tag__String():With Me
.Data("_output")=ReSubmitFields(.Data("string"))
End With:End Function
Public Function tag_pagebreak():tag_pagebreak=True:With Me
If Query.EXT="PDF" Or UCase(j("f"))="FEEDS/PDF" Then .Data("_output")="<pagebreak/>"
End With:End Function
Public Function tag_hsp():tag_hsp=True:With Me
.Data("_output")=hsp(.Data("args"))
End With:End Function
Public Function tag_vsp():tag_vsp=True:With Me
.Data("_output")=vsp(.Data("args"))
End With:End Function
Public Function tag_GetResponseOnLoad():tag_GetResponseOnLoad=tag_GetResponseLazy:End Function
Public Function tag_GetResponseLazy():tag_GetResponseLazy=True:Call tag__String():With Me
If .Data("string")="" Then .Data("string")=.Data("url")
If .Data("string")="" Then .Data("string")=.Data("file")
If .Data("string")="" Then .Data("string")=.Data("args")
.Data("string")=rxReplace(.Data("string"),"\<\/?(\w+)\>","")
If .Data("string")<>"" Then
.Data("_output")=GetResponseLazy(.Data("string"))
Err.Clear()
End If
End With:End Function
Public Function tag_GetResponse():tag_GetResponse=tag_http:End Function
Public Function tag_http():tag_http=True:Call tag__String():With Me
If .Data("string")="" Then .Data("string")=.Data("url")
If .Data("string")="" Then .Data("string")=.Data("file")
If .Data("string")="" Then .Data("string")=.Data("args")
.Data("string")=rxReplace(.Data("string"),"\<\/?(\w+)\>","")
If .Data("string")<>"" Then
.Data("_output")=GetResponse(.Data("string"))
Err.Clear()
End If
End With:End Function
Public Function tag_dir():tag_dir=True:Call tag_tot():Call tag__String():With Me
Call tag__DATA_Prepare_Template()
.Data("F")=.TData("string,folder,dir,directory,path,root,base")
.Data("E")=.TData("type,ext,extension,filter,ftr")
DBG "tag_files> Find Files"
Dim rF,o
Set rF=GetFilesOfType(.Data("F"),.Data("E"))
DBG "tag_files> Got Files"
If Exists(rF)Then
.Data("files")=rF.RecordCount
DBG "tag_files> rF.RecordCount="&.Data("files")
o=MergeRS(rF,.Data("template"))
Else
DBG "tag_files> no files"
End If
DBG "tag_files> output "&Len(o)&"bytes"
.Data("_output")=o
End With:End Function
Public Function tag_files():tag_files=tag_dir:End Function
Public Function tag_filelist():tag_filelist=tag_dir:End Function
Public Function tag_file_list():tag_file_list=tag_dir:End Function
Public Function tag_listfiles():tag_listfiles=tag_dir:End Function
Public Function tag_list_files():tag_list_files=tag_dir:End Function
Public Function tag_file():tag_file=True:Call tag__String():Call tag__String():With Me
.Data("string")=.TData("string,url,file")
If .Data("string")<>"" Then
.Data("_output")=ReadFile(.Data("string"))
End If
End With:End Function
Public Function tag_ssi():tag_ssi=tag_file:End Function
Public Function tag_inc():tag_inc=tag_file:End Function
Public Function tag_include():tag_include=tag_file:End Function
Public Function tag_readfile():tag_readfile=tag_file:End Function
Public Function tag_vatno():tag_vatno=True:With Me
On Error Resume Next
Dim iVatNoAcc:iVatNoAcc=VOZ(.Data("args"))
If iVatNoAcc>0 Then
.Data("_output")=shop.Accounts.Trade.FieldByID("vat",iVatNoAcc)
End If
.Data("_output")=shop("vat")
End With:End Function
Public Function tag_optim():tag_optim=tag_optimize:End Function
Public Function tag_optimize():tag_optimize=True
Dim sStr,sKW,sIgnore,oWords,oWord
sStr=TOT(Me.Data("keywords"),Me.Data("args"))
If IsEmpty(sStr)Then
Exit Function
End If
sIgnore="|"&rxStopWords&"|"&rxBadWords&"|search|results|buy|online|shop|"&site.CFG("keywords.negative")&"|"
Dim sUsed:Set sUsed=New QuickString
Dim output:Set output=New QuickString
With output
r.Pattern="\b\w+\b"
Set oWords=r.Execute(sStr)
For Each oWord In oWords
If InStr(1,sIgnore,"|"&oWord.Value&"|")>0 Then
r.Pattern="\b"&oWord.Value&"\b"
sStr=r.Replace(sStr," ")
End If
Next
r.Pattern="([\-\/\\]+)"
sStr=r.Replace(sStr,",")
r.Pattern="([^ ,\w]+)"
sStr=r.Replace(sStr," ")
r.Pattern="\ +"
sStr=r.Replace(sStr," ")
For Each sKW In Split(sStr,",")
sKW=TOE(Trim(sKW&""))
If(sKW<>"")Then
If Not sUsed.Has(sKW)Then
If .IsSet()Then .Add " | "
sUsed.Add sKW
.Add "<a href=""/search/"&URLEncode(sKW)&"/"">"&sKW&"</a>"
End If
End If
Next
End With
Me.Data("_output")=output
End Function
Public Function tag_localinfo():tag_localinfo=True
If Query.FQ("-PRINT")<>"" Then Exit Function
Me.Data("infoTooltip")=TOT(Me.Data("infoTooltip"),Me.Data("infoTitle"))
Me.Data("mapTooltip")=TOT(Me.Data("mapTooltip"),Me.Data("mapTitle"))
Me.Data("statsTooltip")=TOT(Me.Data("statsTooltip"),Me.Data("statsTitle"))
Me.Data("infoDescription")=TOT(Me.Data("statsDescription"),"get more information about the neighbourhood, local transport and leisure facilities")
Me.Data("infoTooltip")=TOT(Me.Data("infoTooltip"),"neighbourhood, transport, leisure and more from UpMyStreet.com")
Me.Data("infoLabel")=TOT(Me.Data("infoLabel"),"local info")
Me.Data("mapDescription")=TOT(Me.Data("mapDescription"),"view a map of |postcode|")
Me.Data("mapTooltip")=TOT(Me.Data("mapTooltip"),"map of |postcode| from MultiMap.com")
Me.Data("mapLabel")=TOT(Me.Data("mapLabel"),"map")
Me.Data("statsDescription")=TOT(Me.Data("statsDescription"),"view government statistics about the area")
Me.Data("statsTooltip")=TOT(Me.Data("statsTooltip"),"view government statistics")
Me.Data("statsLabel")=TOT(Me.Data("statsLabel"),"stats")
Me.Data("infoLink")=TOT(site.CFG("tools.stats.link"),"http://www.upmystreet.com/overview/?l1=|postcode|")
Me.Data("mapLink")=TOT(site.CFG("tools.stats.link"),"http://www.multimap.com/map/browse.cgi?pc=|postcode|")
Me.Data("statsLink")=TOT(site.CFG("tools.stats.link"),"http://neighbourhood.statistics.gov.uk/dissemination/NeighbourhoodProfileSearch.do?profileSearch=|postcode|")
Me.Data("divider")=TOT(Me.Data("divider")," &bull; ")
Me.Data("mode")=LCase(Me.Data("mode"))
If Me.Data("mode")="li" Then Me.Data("mode")="list"
If Me.Data("mode")="fulllist" Or Me.Data("mode")="menu" Then
Me.Data("mode")="fulllist"
output=output&"<ul class=""ListMenu |direction| |class|"" style=""|style|"">"
End If
If InStr(1,Me.Data("mode"),"list")>0 Then Me.Data("li")=True
If IsEmpty(Me.Data("info"))Or IsTrue(Me.Data("info"))Then
Me.Data("info")=True
If Me.Data("li")Then output=output&"<li>"
output=output&"<a title=""|infoTooltip|"" target=""_blank"" href=""|infoLink|"">|infoLabel|</a>"
If Me.Data("li")Then output=output&"</li>"&vbCRLf
End If
If IsEmpty(Me.Data("map"))Or IsTrue(Me.Data("map"))Then
Me.Data("map")=True
If IsTrue(Me.Data("info"))And(Not Me.Data("li"))Then output=output&"|divider|"&vbCRLf
If Me.Data("li")Then output=output&"<li>"
output=output&"<a title=""|mapTooltip|"" target=""_blank"" href=""|mapLink|"">|mapLabel|</a>"
If Me.Data("li")Then output=output&"</li>"&vbCRLf
End If
If IsEmpty(Me.Data("stats"))Or IsTrue(Me.Data("stats"))Then
If IsTrue(Me.Data("info"))Or IsTrue(Me.Data("map"))And(Not Me.Data("li"))Then output=output&"|divider|"&vbCRLf
If Me.Data("li")Then output=output&"<li>"
output=output&"<a title=""|statsTooltip|"" target=""_blank"" href=""|statsLink|"">|statsLabel|</a>"
If Me.Data("li")Then output=output&"</li>"&vbCRLf
End If
If Me.Data("mode")="fulllist" Or Me.Data("mode")="menu" Then
output=output&"</ul>"
End If
output=MergeTemplateStatic(output,Me.Data)
Me.Data("_output")=output
End Function
Public Function tag_searches():tag_searches=tag_qs:End Function
Public Function tag_popsearch():tag_popsearch=tag_qs:End Function
Public Function tag_qs():tag_qs=True
Set Source=site.Qs
Me.Data("@K")=Source.Key
Me.Data("cookieName")="q"
Me.Data("filter")=SiteFilter
Me.Data("max")=VOV(Me.Data("max"),5)
Me.Data("divider")=TOT(Me.Data("divider"),"<br/>")
If Me.Data("action")="" Then Me.Data("action")=Me.Data("mode")
If Me.Data("action")="" Then Me.Data("action")="popular"
tag_qs=Eval("tag_qs_"&Me.Data("action")&"")
End Function
Public Function tag_qs_popular():tag_qs_popular=True
If Me.Data("period")<>"" Then
Select Case LCase(Me.Data("period"))
Case "today":Me.Data("filter")=MergeCriteria("datediff(hh, [last_hit], getdate())<24","AND",Me.Data("filter"))
Case "yesterday":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())=1","AND",Me.Data("filter"))
Case "week":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())<=7","AND",Me.Data("filter"))
Case "month":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())<=31","AND",Me.Data("filter"))
Case "year":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())<=365","AND",Me.Data("filter"))
Case Else
End Select
End If
Set rQs=Source.Run("SELECT TOP "&VOV(Me.Data("max"),5)&" [name],[hits] FROM %TABLE %NOLOCK WHERE "&Me.Data("filter")&" ORDER BY [hits] DESC")
If Not ExistsOrKill(rQs)Then Exit Function
Me.Data("_output")=tag_qs_list(rQs)
End Function
Public Function tag_qs_latest():tag_qs_latest=True
Set rQs=Source.Run("SELECT TOP "&VOV(Me.Data("max"),5)&" [name] FROM %TABLE %NOLOCK WHERE "&Me.Data("filter")&" ORDER BY [last_hit] DESC")
If Not ExistsOrKill(rQs)Then Exit Function
Me.Data("_output")=tag_qs_list(rQs)
End Function
Public Function tag_qs_list(ByRef rQs)
If Not ExistsOrKill(rQs)Then Exit Function
Dim output:Set output=New QuickString
With output
Do While Exists(rQs)
If .IsSet()Then .Add Me.Data("divider")
.Add "<a href=""/search/"&URLEncode(rQs("name"))&"/"" title="""&rQs("name")&" - search results"">"&Slice(rQs("name"),15)&"</a>"
rQs.MoveNext
Loop
End With
tag_qs_list=output
End Function
Public Function tag_refqs():tag_refqs=True
Set Source=site.RefQs
Me.Data("@K")=Source.Key
Me.Data("cookieName")="refq"
Me.Data("filter")=SiteFilter
Me.Data("max")=VOV(Me.Data("max"),5)
Me.Data("divider")=TOT(Me.Data("divider"),"<br/>")
If Me.Data("action")="" Then Me.Data("action")=Me.Data("mode")
If Me.Data("action")="" Then Me.Data("action")="popular"
tag_refqs=Eval("tag_refqs_"&Me.Data("action")&"")
End Function
Public Function tag_refqs_popular():tag_refqs_popular=True
If Me.Data("period")<>"" Then
Select Case LCase(Me.Data("period"))
Case "today":Me.Data("filter")=MergeCriteria("datediff(hh, [last_hit], getdate())<24","AND",Me.Data("filter"))
Case "yesterday":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())=1","AND",Me.Data("filter"))
Case "week":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())<=7","AND",Me.Data("filter"))
Case "month":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())<=31","AND",Me.Data("filter"))
Case "year":Me.Data("filter")=MergeCriteria("datediff(d, [last_hit], getdate())<=365","AND",Me.Data("filter"))
Case Else
End Select
End If
Set rQs=Source.Run("SELECT TOP "&VOV(Me.Data("max"),5)&" [name],[hits] FROM %TABLE %NOLOCK WHERE "&Me.Data("filter")&" ORDER BY [hits] DESC")
If Not ExistsOrKill(rQs)Then Exit Function
Me.Data("_output")=tag_refqs_list(rQs)
End Function
Public Function tag_refqs_latest():tag_refqs_latest=True
Set rQs=Source.Run("SELECT TOP "&VOV(Me.Data("max"),5)&" [name] FROM %TABLE %NOLOCK WHERE "&Me.Data("filter")&" ORDER BY [last_hit] DESC")
If Not ExistsOrKill(rQs)Then Exit Function
Me.Data("_output")=tag_refqs_list(rQs)
End Function
Public Function tag_refqs_list(ByRef rQs)
If Not ExistsOrKill(rQs)Then Exit Function
Dim output:Set output=New QuickString
With output
Do While Exists(rQs)
If .IsSet()Then .Add Me.Data("divider")
.Add "<a href=""/search/"&URLEncode(rQs("name"))&"/"" title="""&rQs("name")&" - search results"">"&Slice(rQs("name"),15)&"</a>"
rQs.MoveNext
Loop
End With
tag_refqs_list=output
End Function
Public Function tag_bookmark():tag_bookmark=True
Dim o:Set o=New QuickString
If Me.Data("link")="" Then Me.Data("link")=j("link")
If Me.Data("name")="" Then Me.Data("name")=Me.Data("title")
If Me.Data("name")="" Then Me.Data("name")=j("name")
If Me.Data("name")="" Then Me.Data("name")=j("title")
If Me.Data("label")="" Then
Me.Data("label")="<span>&nbsp;|engineName|</span>"
Me.Data("label")="<span>"&"&nbsp;"&"|engineName|"&"</span>"&""
End If
If Me.Data("template")="" Then
Me.Data("template")="<li>"&"<a rel=""nofollow"" target=""_blank"" href=""|engineLink|"""&" title=""|engineName|:|T|"""&">"&"<img height='16' width='16' align='bottom'"&" src=""/@/share/|engine|.gif"""&" alt=""|engineName|:|T|"""&"/>"&"<span>&nbsp;|label|</span>"&"</a>"&"</li>"&""
End If
If Me.Data("only")="" Then Me.Data("only")=Me.Data("filter")
Me.Data("only")=Replace(Me.Data("only"),",","|")
Me.Data("L")=Me.Data("link")
Me.Data("T")=Me.Data("name")
Me.Data("TE")=UrlEncode(Me.Data("name"))
Dim engs,eng
Set engs=tag_bookmark_engines
For Each eng In engs
If Me.Data("only")<>"" Then
If Not rxTest(eng,Me.Data("only"))Then eng=""
End If
If Me.Data("not")<>"" Then
If rxTest(eng,Me.Data("not"))Then eng=""
End If
If eng<>"" Then
Me.Data("engine")=eng
eng=Split(engs(eng)," % ")
Me.Data("engineName")=eng(0)
Me.Data("engineLink")=eng(1)
o.Add mergeTemplate(Me.Data("template"),Me.Data)
End If
Next
Me.Data("_output")=o
End Function
Public Function tag_bookmark_engines()
Dim d:Set d=Dictionary()
d("digg")="Digg % http://digg.com/submit?phase=2&amp;url=|L|&amp;title=|TE|"
d("facebook")="Facebook % javascript:var%20d=document,f='http://www.facebook.com/share',l=d.location,e=encodeURIComponent,p='.php?bm&v=3&u='+e('|L|')+'&t='+e('|T|');try{if(!/^(.*\.)?facebook\.[^.]*$/.test(l.host))throw(0);share_internal_bookmarklet()}catch(z){a=function(){if(!window.open(f+'r'+p,'sharer','toolbar=0,status=0,resizable=0,width=626,height=436'))l.href=f+p};if(/Firefox/.test(navigator.userAgent))setTimeout(a,0);else{a()}}void(0)"
d("delicious")="Del.icio.us % http://del.icio.us/post?url=|L|&amp;title=|TE|"
d("google")="Google Bookmarks % http://www.google.com/bookmarks/mark?op=edit&amp;bkmk=|L|&amp;title=|TE|"
d("yahoo")="Yahoo! My Web % http://myweb2.search.yahoo.com/myresults/bookmarklet?u=|L|&amp;t=|TE|"
d("windows")="Windows Live % https://favorites.live.com/quickadd.aspx?marklet=1&amp;mkt=en-us&amp;url=|L|&amp;top=1&amp;title=|TE|"
d("technorati")="Technorati % http://www.technorati.com/faves?add=|L|&amp;title=|TE|"
d("magnolia")="ma.gnolia % http://ma.gnolia.com/bookmarklet/add?url=|L|&amp;title=|TE|"
d("spurl")="Spurl % http://www.spurl.net/spurl.php?title=|TE|&amp;url=|TE|"
d("furl")="Furl % http://www.furl.net/storeIt.jsp?u=|L|&amp;t=|TE|"
d("reddit")="Reddit % http://reddit.com/submit?url=|L|&amp;title=|TE|"
d("netscape")="Netscape % http://www.netscape.com/submit/?U=|L|&amp;T=|TE|"
d("citeUlike")="CiteULike % http://www.citeulike.org/posturl?url=|L|&amp;title=|TE|"
d("wists")="Wists % http://wists.com/r.php?c=&amp;r=|L|&amp;title=|TE|"
d("simpy")="Simpy % http://www.simpy.com/simpy/LinkAdd.do?href=|L|&amp;title=|TE|"
d("newsvine")="Newsvine % http://www.newsvine.com/_tools/seed&amp;save?u=|L|&amp;h=|TE|"
d("blinklist")="Blinklist % http://www.blinklist.com/index.php?Action=Blink/addblink.php&amp;Description=&amp;Url=|L|&amp;Title=|TE|"
d("fark")="Fark % http://cgi.fark.com/cgi/fark/edit.pl?new_url=|L|&amp;linktype=&amp;new_comment=|TE|"
d("blogmarks")="Blogmarks % http://blogmarks.net/my/new.php?mini=1&amp;simple=1&amp;url=|L|&amp;title=|TE|"
d("smarking")="Smarking % http://smarking.com/editbookmark/?url=|L|&amp;title=|TE|"
d("tailrank")="Tailrank % http://tailrank.com/share/?link_href=|L|&amp;title=|TE|"
Set tag_bookmark_engines=d
End Function
Public Function tag_exe():tag_exe=tag_exec:End Function
Public Function tag_exec():tag_exec=True
Call tag__String()
If IsTest Then DBG.Warning "##EXE-TAG:"&Me.Data("string")
Me.Data("_output")=EXE(Me.Data("string"))
End Function
Public Function tag_iifcaptcha():tag_iifcaptcha=tag_ifcaptcha:End Function
Public Function tag_if_captcha():tag_if_captcha=tag_ifcaptcha:End Function
Public Function tag_ifcaptcha():tag_ifcaptcha=True:With Me
.Data("string")=UseCaptcha
Call tag__iif()
End With:End Function
Public Function tag_captcha():tag_captcha=True:With Me
.Data("_output")=reCAPTCHA_Write
End With:End Function
Public Function tag_exe2():tag_exe2=tag_exec2:End Function
Public Function tag_exec2():tag_exec2=True
On Error Resume Next
Call tag__String()
Me.Data("_output")=EXE2(Me.Data("string"))
End Function
Function Switch():Switch=True
Dim output:Set output=New QuickString
With output
If IsEmpty(Me.Data("field"))Then Exit Function
If IsEmpty(Me.Data("label"))Then
Me.Data("label")=Me.Data("field")
End If
If IsEmpty(Me.Data("yesLabel"))Then
Me.Data("yesLabel")=Me.Data("label")
End If
If IsEmpty(Me.Data("noLabel"))Then
Me.Data("noLabel")="not "&Me.Data("label")
End If
Me.Data("currentValue")=Decode(Query.FQ(Me.Data("field")))
If IsSet(Me.Data("currentValue"))And Me.Data("currentValue")<>"-" Then
If IsTrue(Me.Data("currentValue"))Then
.Add "<a href="""&ReSubmitQuery(TOT(Me.Data("field"),"ERROR")&"=no")&""">"&Me.Data("yesLabel")&"</a>"&vbCRLf
Else
.Add "<a href="""&ReSubmitQuery(TOT(Me.Data("field"),"ERROR")&"=yes")&""">"&Me.Data("noLabel")&"</a>"&vbCRLf
End If
.Add "<a href="""&ReSubmitQuery("~"&TOT(Me.Data("field"),"ERROR")&"")&"""><img src="""&App.Image("menu/no")&""" height=""8""></a>"
.Add TOT(Me.Data("after"),"[hsp 10]")&vbCRLf
End If
End With
Me.Data("_output")=output
End Function
Public Function tag_pdf():tag_pdf=True
Call tag__String()
Dim action,output
action=Me.TData("string")
If rxTest(action,"^\w+$")Then
If rxTest(action,"(page)?break")Then
output="<pagebreak/>"
End If
Else
End If
If output<>"" Then
output="{Is.FEEDS.PDF}"&output&"{/Is.FEEDS.PDF}"
End If
Me.Data("_output")=output
End Function
Public Function tag_brands():tag_brands=tag__folders("brand"):End Function
Public Function tag_folders():tag_folders=tag__folders("folder"):End Function
Public Function tag_places():tag_places=tag__folders("place"):End Function
Public Function tag_categories():tag_categories=tag__folders("category"):End Function
Public Function tag_ranges():tag_ranges=tag__folders("range"):End Function
Public Function tag_prices():tag_prices=tag_ranges:End Function
Public Function tag_price_ranges():tag_price_ranges=tag_ranges:End Function
Public Function tag__folders(ByVal sID):tag__folders=True
Dim rsP
Me.Data("sql")="SELECT ["&sID&"Name] AS 'name', COUNT(%KEY) AS n FROM %TABLE %NOLOCK WHERE %SEARCH GROUP BY ["&sID&"Name] ORDER BY ["&sID&"Name] ASC "
Set rsP=Source.Run(Me.Data("sql"))
If Not ExistsOrKill(rsP)Then Exit Function
If Me.Data("orientation")="" Then Me.Data("orientation")=Me.Data("direction")
If Me.Data("orientation")="" Then Me.Data("direction")="Vertical"
If Me.Data("class")<>"" Then Me.Data("class")=" "&Me.Data("class")
Dim output:Set output=New QuickString
With output
.Add "<ul class=""ListMenu "&Me.Data("orientation")&Me.Data("class")&""">"
Do While Exists(rsP)
.Add "<li>"
.Add "<a href="""&j("link")&"-"&sID&"Name="&UrlEncode(TOT(rsP("name"),"-"))&"/?q="&Query.FQ("q")&""">"
.Add TOT(rsP("name"),"Unknown")&" ("&rsP("n")&")"
.Add "</a>"
.Add "</li>"
rsP.MoveNext
Loop
.Add "</ul>"
End With
Me.Data("_output")=output
End Function
Public Function tag_asset():tag_asset=tag_assets:End Function
Public Function tag_assets():tag_assets=True
Dim output:Set output=New QuickString
If Not(Me.Data("asset")<>"")Then
Me.Data("asset")=TOE(Me.Data("args"))
End If
If Me.Data("asset")<>"" Then
If Me.Data("id")<>"" Then
Query.Data(Mid(Me.Data("asset"),1,Len(Me.Data("asset"))-1)&"")=VOZ(Me.Data("id"))
End If
Select Case LCase(Me.Data("asset"))
Case "downloads":output.Add FE__ASSETS_Downloads
Case "images":output.Add FE__ASSETS_Images
Case "links":output.Add FE__ASSETS_Links
Case "reviews":output.Add FE__ASSETS_Reviews
Case "faqs":output.Add FE__ASSETS_FAQs
Case "chapters":output.Add FE__ASSETS_Chapters
Case Else
output.Add "Undefined asset: '"&Me.Data("asset")&"'"
End Select
Else
output.Add "ERROR: asset undefined: '"&Me.Data("asset")&"'"
End If
Me.Data("_output")=output
End Function
Public Function tag_inventory():tag_inventory=tag_product:End Function
Public Function tag_products():tag_products=tag_product:End Function
Public Function tag_product():tag_product=True
Set Source=invy.Products
With Source
Me.Data("@K")=.Key
Me.Data("cookieName")=.ClassName
.ID="":.ParentID=""
End With
tag_product=tag_list()
End Function
Public Function tag_product_brands():tag_product_brands=tag__folders("brand"):End Function
Public Function tag_product_folders():tag_product_folders=tag__folders("folder"):End Function
Public Function tag_product_categories():tag_product_categories=tag__folders("category"):End Function
Public Function tag_product_ranges():tag_product_ranges=tag__folders("range"):End Function
Public Function tag_product_price_ranges():tag_product_price_ranges=tag_product_ranges:End Function
Public Function tag_news():tag_news=tag_article:End Function
Public Function tag_articles():tag_articles=tag_article:End Function
Public Function tag_article()
Set Source=site.Articles
With Source
Me.Data("@K")=.Key
Me.Data("cookieName")=.ClassName
.ID="":.ParentID=""
End With
tag_article=tag_list()
End Function
Public Function tag_article_folders():tag_article_folders=tag__folders("folder"):End Function
Public Function tag_article_categories():tag_article_categories=tag__folders("category"):End Function
Public Function tag_article_places():tag_article_places=tag__folders("place"):End Function
Public Function tag_article_locations():tag_article_locations=tag_article_places:End Function
Public Function tag_eman():tag_eman=tag_event:End Function
Public Function tag_events():tag_events=tag_event:End Function
Public Function tag_event()
Set Source=eman.Events
With Source
Me.Data("@K")=.Key
Me.Data("cookieName")=.ClassName
.ID="":.ParentID=""
End With
tag_event=tag_list()
End Function
Public Function tag_event_folders():tag_event_folders=tag__folders("folder"):End Function
Public Function tag_event_categories():tag_event_categories=tag__folders("category"):End Function
Public Function tag_event_places():tag_event_places=tag__folders("place"):End Function
Public Function tag_event_locations():tag_event_locations=tag_event_places:End Function
Public Function tag_rees():tag_rees=tag_property:End Function
Public Function tag_properties():tag_properties=tag_property:End Function
Public Function tag_property()
Set Source=rees.Properties
With Source
Me.Data("@K")=.Key
Me.Data("cookieName")=.ClassName
.ID="":.ParentID=""
End With
tag_property=tag_list()
End Function
Public Function tag_property_folders():tag_property_folders=tag__folders("folder"):End Function
Public Function tag_property_categories():tag_property_categories=tag__folders("category"):End Function
Public Function tag_property_places():tag_property_places=tag__folders("place"):End Function
Public Function tag_property_locations():tag_property_locations=tag_property_places:End Function
Public Function tag_property_ranges():tag_property_ranges=tag__folders("range"):End Function
Public Function tag_property_price_ranges():tag_property_price_ranges=tag_property_ranges:End Function
Public Function tag_ModuleOutput()
tag_ModuleOutput=True
Me.Data("template")="content"
Call tag_GetTemplate()
End Function
Public Function tag_res():tag_res=tag_resolution:End Function
Public Function tag_resolution():tag_resolution=True
Me.Data("_output")=TOT(MyCookies.Cookie("Res"),"LRes")
End Function
Public Function tag_browser():tag_browser=tag_browser_tag:End Function
Public Function tag_browsertag():tag_browsertag=tag_browser_tag:End Function
Public Function tag_browser_tag():tag_browser_tag=True
If IsIE Then Me.Data("_output")="IE"
If IsFF Then Me.Data("_output")="FF"
End Function
Public Function tag_merge():tag_merge=True
Dim o
Call tag__DATA_Prepare_Template()
o=MergeTemplateOnce(Me.Data("template"),Me.Data)
Me.Data("_output")=o
End Function
Public Function tag_tmpl():tag_tmpl=tag_get_template:End Function
Public Function tag_template():tag_template=tag_get_template:End Function
Public Function tag_gettemplate():tag_gettemplate=tag_get_template:End Function
Public Function tag_get_template():tag_get_template=True
Call tag__String()
Dim o,s,t,f
Dim tried
Set tried=New QuickString
Dim system_tmpl:system_tmpl=IIf(IsFront,"FE__TMPL","BO__TMPL")
Dim custom_tmpl:custom_tmpl=TOT(j("t"),system_tmpl)
If Me.TData("in-ram,from-ram")<>"" Then
custom_tmpl=custom_tmpl&"_FromRAM"
system_tmpl=system_tmpl&"_FromRAM"
End If
DBG "TAG: getTemplate - t='"&j("t")&"'"
DBG "TAG: getTemplate - system_tmpl='"&system_tmpl&"'"
DBG "TAG: getTemplate - custom_tmpl='"&custom_tmpl&"'"
For Each s In Split("file,template,string,a,b,c,d,e,path",",")
t=Me.Data(s)
If o="" And t<>"" Then
If InStr(1,tried,"("&t&")")>0 Then
Else
tried.Add "("&t&")"
f=custom_tmpl
If InStr(1,"!@/",Left(t,1))Then
f=system_tmpl
End If
If InStr(1,t,":")>0 Then
f=Mid(t,1,InStr(1,t,":")-1)
t=Mid(t,InStr(1,t,":")+1)
End If
DBG "TAG: getTemplate - "&f&"("""&t&""")"&"..."
Select Case UCase(f)
Case "FE__TMPL":o=FE__TMPL(t)
Case "BO__TMPL":o=BO__TMPL(t)
Case "BO_FORM__TMPL":o=BO_FORM__TMPL(t)
Case Else:o=EXE(f&"("""&t&""")")
End Select
DBG "TAG: getTemplate - "&f&"("""&t&""")"&" >>> "&Len(o)&" bytes"
If o<>"" Then
If Left(t,1)="\" Then
Me.Data("escape")="y"
End If
If Me.Data("escape")<>"" Then
o=rxReplace(o,"(""|'|\\)","\$1")
End If
End If
End If
End If
Next
Me.Data("_output")=o
End Function
Public Function tag_tmplfile():tag_tmplfile=tag_get_template_file:End Function
Public Function tag_templatefile():tag_templatefile=tag_get_template_file:End Function
Public Function tag_gettemplatefile():tag_gettemplatefile=tag_get_template_file:End Function
Public Function tag_get_template_file():tag_get_template_file=True
Call tag__String()
Dim o,s,t
Dim tried
Set tried=New QuickString
For Each s In Split("file,template,string,a,b,c,d,e,path",",")
t=Me.Data(s)
If o="" And t<>"" Then
If InStr(1,tried,"("&t&")")>0 Then
Else
tried.Add "("&t&")"
o=TemplateFile(t)
End If
End If
Next
Me.Data("_output")=o
End Function
End Class
Function FWX_ParsedData(ByVal f)
Dim k
k=Left(f,1)
If k="=" And Len(f)>1 Then
f=Mid(f,2)
If rxTest(f,"^\W")Then
k=Left(f,1)
f=Mid(f,2)
If InStr(1,f,"~")>0 Then
f=parser.Merge(k&f&k)
Else
Select Case k
Case "#":f=j(f)
Case "$":f=App.CFG(f)
Case "?":f=Query.FQ(f)
Case "!":f=TOT(MySession(f),MyCookies(f))
Case ":":f=GetTemplate(f)
Case "|":
If Exists(j("rs"))Then
f=RSFieldGet(j("rs"),f)
Else
f="ERROR 728 ("&f&")"
End If
Case Else:f="MergeError7("&f&")"
End Select
End If
ElseIf rxTest(f,"^\w+\:")Then
k=rxSelect(f,"^\w+")
f=rxReplace(f,"^\w+\:","")
Select Case LCase(k)
Case "c","cookie","cookies":f=MyCookies(f)
Case "s","session":f=MySession(f)
Case "t","tmpl","template":f=GetTemplate(f)
Case "f","fso","file":f=ReadFile(f)
Case Else:
DBG.Show=True
f="ERROR 627: dynamic merging failed for "&k&"("&f&")"
End Select
Else
If f<>"" Then
f=EXE(f)
End If
End If
ElseIf k="|" Then
If rxTest(f,"^\|.+\|$")Then
End If
End If
FWX_ParsedData=f
End Function
Const fwx_bitly_api="120b320017fe1abd45663c4354805e47b5fc33bb"
Function bitly(longURL)
Dim req,url
url="https://api-ssl.bitly.com/v3/shorten?domain=j.mp&format=txt&access_token="&fwx_bitly_api&"&longUrl="&longURL&"&"
set req=Server.CreateObject("MSXML2.ServerXMLHTTP")
req.open "GET",url,false
req.send ""
bitly=req.responseText
set req=nothing
END FUNCTION
Function ShortenUrl(ByVal s)
ShortenUrl=bitly(s)
End Function
Function ShortUrl(ByVal s)
ShortUrl=bitly(s)
End Function
Const TinyURL_Server="http://tinyurl.com/api-create.php"
Function TinyURL(ByVal s)
Dim u,x
u=TinyURL_Server&"?url="&Server.URLEncode(s)&""
x=GetResponse(u)
If x<>"" Then
s=x
End If
TinyURL=s
End Function
Function CloudFlare_DevMode_ON()
Exit Function
DBG "CloudFlare_DevMode_ON"
CloudFlare_DevMode_ON=CloudFlare_DevMode(1)
End Function
Function CloudFlare_DevMode_OFF()
DBG "CloudFlare_DevMode_OFF"
CloudFlare_DevMode_OFF=CloudFlare_DevMode(0)
End Function
Function CloudFlare_DevMode(ByVal v)
DBG "CloudFlare_DevMode"
CloudFlare_DevMode=CloudFlareAPI("devmode",v)
End Function
Function CloudFlarePurge()
DBG "CloudFlarePurge"
CloudFlarePurge=CloudFlareAPI("fpurge_ts","1")
End Function
Function CloudFlareAPI(ByVal a,ByVal v)
DBG "CloudFlareAPI('"&a&"','"&v&"')..."
If IsTrue(IsStaff)Or IsTrue(IsStaffPC)Then
Else
DBG "CloudFlareAPI('"&a&"','"&v&"') must be logged in (leaving)"
CloudFlareAPI="command unavailable unless authenticated"
Exit Function
End If
Dim token,uri,qry,res,api
If Not IsObject(site)Then
Set site=New fwxSITE
Set app=New fwxAPP
End If
If IsObject(site)Then
DBG "CloudFlareAPI('"&a&"','"&v&"') get settings in site"
If token="" Then token=site.CFG("cloudflare.token")
If email="" Then email=site.CFG("cloudflare.email")
End If
If IsObject(FWX)Then
DBG "CloudFlareAPI('"&a&"','"&v&"') get settings in FWX.site? ('"&email&"'/'"&token&"')"
If token="" Then token=FWX.Setting("site.cloudflare.token")
If email="" Then email=FWX.Setting("site.cloudflare.email")
DBG "CloudFlareAPI('"&a&"','"&v&"') get settings in FWX? ('"&email&"'/'"&token&"')"
If token="" Then token=FWX.Setting("cloudflare.token")
If email="" Then email=FWX.Setting("cloudflare.email")
End If
DBG "CloudFlareAPI('"&a&"','"&v&"') get settings in Application(site) ('"&email&"'/'"&token&"')"
If email="" Then email=Application.Contents("site.cloudflare.email")
If token="" Then token=Application.Contents("site.cloudflare.token")
DBG "CloudFlareAPI('"&a&"','"&v&"') get settings in Application() ('"&email&"'/'"&token&"')"
If email="" Then email=Application.Contents("cloudflare.email")
If token="" Then token=Application.Contents("cloudflare.token")
DBG "CloudFlareAPI('"&a&"','"&v&"') email='"&email&"'?"
DBG "CloudFlareAPI('"&a&"','"&v&"') token='"&token&"'?"
If email<>"" And token<>"" Then
uri="https://www.cloudflare.com/api_json.html"
qry="tkn="&token&"&email="&email&"&z="&MyDomain&"&a="&a&"&v="&URLEncode(v)&"&"
api=uri&"?"&qry
res=GetResponse(api)
End If
DBG "CloudFlareAPI('"&a&"','"&v&"') uri='"&Encode(uri)&"'"
DBG "CloudFlareAPI('"&a&"','"&v&"') qry='"&Encode(qry)&"'"
DBG "CloudFlareAPI('"&a&"','"&v&"') res='"&Encode(res)&"'"
If res<>"" Then
If rxTest(res,"\{")Then
Dim o,d,msg
Set d=JSON2DIC(JSON.Parse(res))
If msg="" Then msg=d("response.fb_msg")
If msg="" Then msg=d("response.msg")
If msg="" Then msg=d("msg")
If msg="" Then msg=d("result")
If msg="" Then msg=res
Set d=Nothing
End If
End If
CloudFlareAPI=msg
End Function
Function GeoCode(ByVal q)
Set GeoCode=LoadXML("http://maps.google.com/maps/geo?output=xml&q="&q&"&")
End Function
Const GoogleMapsKey__localhost="AIzaSyDh2waeOQ68ZrQuprr82X3vCH39YVHrLAE"
Function GoogleMapsKey()
Dim k:k=j("GoogleMapsKey")
If k="" Then k=GoogleAPIKey_Client()
If k="" Then k=App.CFG("google.maps.key."&MyServerName)
If k="" Then k=App.CFG("google.maps.key."&MyDomainForCode)
If k="" Then k=App.CFG("google.maps.key."&MyDomain)
If k="" Then k=App.CFG("google.maps.key."&MyDomainID)
If k="" Then k=Site.CFG("google.maps.key")
If k="" Then k=App.CFG("google.maps.key")
j("GoogleMapsKey")=k
GoogleMapsKey=k
End Function
Function GoogleAPIPost(PostUrl,sPostData)
Dim XmlHttp,BasicAuthentication,ResponseData,key
key=GoogleAPIKey_Server()
If key="" Then
DBG.Warning "Google API Key Missing"
Else
Set XmlHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
XmlHttp.Open "POST",PostUrl,false
XmlHttp.SetRequestHeader "Authorization",key
XmlHttp.Send sPostData
ResponseData=XmlHttp.ResponseText
GoogleAPIPost=ResponseData
Set XmlHttp=Nothing
If InStr(1,ResponseData,"error")>0 Then
DBG.Warning "Google API Failed: "&key
End If
End If
End Function
Function GoogleAPIKey()
GoogleAPIKey=GoogleAPIKey_Client()
End Function
Function GoogleAPIKey_Client()
Dim k:k=j("GoogleAPIKey_Client")
If k="" Then k=App.CFG("google.api.client."&MyServerName)
If k="" Then k=App.CFG("google.api.client."&MyDomainForCode)
If k="" Then k=App.CFG("google.api.client."&MyDomain)
If k="" Then k=App.CFG("google.api.client."&MyDomainID)
If k="" Then k=Site.CFG("google.api.client")
If k="" Then k=App.CFG("google.api.client")
j("GoogleAPIKey_Client")=k
GoogleAPIKey_Client=k
End Function
Function GoogleAPIKey_Server()
Dim k:k=j("GoogleAPIKey_Server")
If k="" And IsObject(App)Then k=App.CFG("google.api.server."&MyServerName)
If k="" And IsObject(App)Then k=App.CFG("google.api.server."&MyDomainForCode)
If k="" And IsObject(App)Then k=App.CFG("google.api.server."&MyDomain)
If k="" And IsObject(App)Then k=App.CFG("google.api.server."&MyDomainID)
If k="" And IsObject(Site)Then k=Site.CFG("google.api.server")
If k="" And IsObject(App)Then k=App.CFG("google.api.server")
If k="" And IsObject(FWX)Then k=FWX.Setting("google.api.server")
j("GoogleAPIKey_Server")=k
GoogleAPIKey_Server=k
End Function
Function GoogleLocalLink(ByVal bID)
GoogleLocalLink="http://maps.google.co.uk/maps?latlng="&bID&""
End Function
Function GoogleMapsAPI()
GoogleMapsAPI="<script src=""https://maps.googleapis.com/maps/api/js?key="&GoogleMapsKey()&""" type=""text/javascript""></script>"&vbCRLf
End Function
Function GoogleGeocodeAPI(ByVal q)
Dim o
o=j("GoogleGeocodeAPI:"&q)
If o="" Then
If rxTest(q,"^[\.\-\,0-9]+,\s?[\.\-\,0-9]+$")Then
o=q
ElseIf LCase(q)="london" Then
o=COORDS_London
Else
o=GetResponse("https://maps.googleapis.com/maps/api/geocode/json?sensor=false&address="&q&"")
End If
j("GoogleGeocodeAPI:"&q)=TOT(o," ")
End If
DBG "GOOGLE > GoogleGeocodeAPI > q='"&q&"' > c='"&c&"' > RESULT: "&o
GoogleGeocodeAPI=o
End Function
Const COORDS_London="51.5073509,-0.1277583"
Function GoogleTimezoneAPI(ByVal q)
Dim l:l=TOT(q,COORDS_London)
Dim c:c=GoogleGeocodeAPI(l)
Dim k:k="GoogleTimezoneAPI:"&c
Dim u:u="https://maps.googleapis.com/maps/api/timezone/json?sensor=false&location="&c&"&timestamp="&DateDiff("s","1 Jan 1970 00:00",UTCNow())&"&key="&GoogleAPIKey_Server()&"&"
Dim o
o=j(k)
If o="" Then
o=Application.Contents(k)
If o="" Then
DBG "GOOGLE > GoogleTimezoneAPI > q='"&q&"' > c='"&c&"' > QUERY API: "&u
o=GetResponse(u)
DBG "GOOGLE > GoogleTimezoneAPI > q='"&q&"' > c='"&c&"' > API RESPONSE: "&o
End If
j(k)=TOT(o," ")
End If
o=TOE(o)
DBG "GOOGLE > GoogleTimezoneAPI > q='"&q&"' > c='"&c&"' > RESULT: "&o
GoogleTimezoneAPI=o
End Function
Function GoogleTimezoneIn(ByVal q)
Dim x
x=GoogleTimezoneAPI(q)
x=rxReplace(x,"[\t\r\n\s]+","")
x=rxSelect(x,"""timeZoneId"":""[^""]+")
x=rxSelect(x,"[^""]+$")
GoogleTimezoneIn=x
End Function
Function GoogleTimezoneOffset(ByVal q)
Dim x
x=GoogleTimezoneAPI(q)
x=rxReplace(x,"[\t\r\n\s]+","")
x=rxSelect(x,"""dstOffset"":\d+")
x=rxSelect(x,"\d+$")
GoogleTimezoneOffset=x
End Function
Function GoogleGeocodeRIP(ByVal q)
j("GoogleGeocodeRIP:"&q)=GetHumanResponse("http://maps.google.co.uk/maps?f=q&oe=ISO-8859-1&q="&q&"")
GoogleGeocodeAPI=j("GoogleGeocodeRIP:"&q)
End Function
Function GoogleMaps(ByVal q)
GoogleMaps=GoogleGeocodeRIP(q)
End Function
CONST rxCOORDS="[\-\.\d]+,[\-\.\d]+"
Function IsCoords(ByVal s)
IsCoords=rxTest(s,rxCOORDS)
End Function
Function GetCoords(ByVal q)
GetCoords=GoogleCoords(q)
End Function
Function GoogleCoords(ByVal q)
GoogleCoords=GoogleGeocode(q)
End Function
Function GoogleGeocode(ByVal q)
Dim s,coords,lat,lng,e,ex
Dim cache,cache_fail,cache_file
cache=True
cache_file="/coords/"&FileName(q)&".txt"
q=TOE(q)
If q="" Then
Exit Function
End If
If IsCoords(q)Then
coords=q
End If
If cache And coords="" Then
coords=Application.Contents("GoogleGeocode:"&UCase(q))
End If
If coords<>"" Then
If coords=cache_fail Then
coords=""
ElseIf IsCoords(coords)Then
GoogleGeocode=coords
End If
Exit Function
End If
s=GoogleGeocodeAPI(q)
If s<>"" Then
DBG "GOOGLE > GoogleGeocode > API Response: "&s
s=rxReplace(s,"[\s\t\n\r]","")
e=rxReplace(s,"^\{""results"":\[\],""status"":""([^""]*)""\}$","$1")
If e<>"" And e<>s Then
Select Case e
Case "ZERO_RESULTS":ex="addess not found '"&q&"'"
Case "REQUEST_DENIED":ex="request was malformed"
Case "OVER_QUERY_LIMIT":ex="daily quota exceeded"
End Select
Exit Function
End If
s=rxSelect(s,"""location""\:\{[^\}]+\}")
s=rxReplace(s,"[^\,\.0-9\-]","")
coords=s
lat=rxSelect(coords,"\,\s*[0-9\-\.]+"):lat=rxSelect(lat,"[0-9\-\.]+")
lng=rxSelect(coords,"\[\s*[0-9\-\.]+"):lng=rxSelect(lng,"[0-9\-\.]+")
Else
s=GoogleGeocodeRIP(q)
coords=rxSelect(s,"center\:\s*\{lat\:\s*([0-9\-\.]+)\,\lng\:\s*[0-9\-\.]+\}")
lat=rxSelect(coords,"lat\:\s*[0-9\-\.]+"):lat=rxSelect(lat,"[0-9\-\.]+")
lng=rxSelect(coords,"lng\:\s*[0-9\-\.]+"):lng=rxSelect(lng,"[0-9\-\.]+")
End If
If lat<>"" And lng<>"" Then
coords=lat&","&lng
DBG "GOOGLE > GoogleGeocode > COORDS: "&coords
Else
DBG "GOOGLE > GoogleGeocode > FAILED!"
End If
GoogleGeocode=coords
coords=TOT(coords,"*")
Application.Contents("GoogleGeocode:"&UCase(q))=coords
WriteToFile cache_file,coords
End Function
Function GetDistance(ByVal a,ByVal b)
Dim s,x,d
a=GetCoords(a)
b=GetCoords(b)
If a="" Or b="" Then
Else
s="SELECT dbo.[fncApp_Tool_Distance_LatLng]("&a&","&b&") distance"
Set x=app.Run(s)
If Exists(x)Then d=VOZ(x("distance"))
Set x=Nothing
End If
GetDistance=d
End Function
Function BloggerFeed(ByVal s)
s=IIf(InStr(1,s,"/")>0,s,"http://"&s&".blogspot.com/feeds/posts/default")
s=rxReplace(s,"(alt)=(\w+)&?","")
s=s&IIf(rxTest(s,"\?$")Or rxTest(s,"\&$"),"",IIf(rxTest(s,"\?"),"&","?"))
s=s&"alt=rss&"
BloggerFeed=s
End Function
Const PicasaBase="http://picasaweb.google.com/"
CONST PICASA_FEED_ALBUMS="http://picasaweb.google.com/data/feed/base/user/USERNAME?kind=album&alt=rss&"
CONST PICASA_FEED_PHOTOS="http://picasaweb.google.com/data/feed/base/user/USERNAME/album/ALBUMNAME?kind=photo&alt=rss&"
Function PicasaAlbums(ByVal s)
s=Trim(Replace(s,PicasaBase,""))
s=Replace(s,"&amp;","&")
s=rxReplace(s,"(kind|alt|hl)=(\w+)&?","")
q=rxSelect(s,"\?.*$"):s=rxReplace(s,"\?.*$","")
s=rxReplace(s,"^/|/$","")
s=IIf(InStr(1,s,"/")>0,Replace(s,"feed/api","feed/base"),"data/feed/base/user/"&s&"")
If q<>"" Then s=s&q
s=s&IIf(rxTest(s,"\?$")Or rxTest(s,"\&$"),"",IIf(rxTest(s,"\?"),"&","?"))
s=PicasaBase&s&"alt=rss&kind=album&"
PicasaAlbums=s
End Function
Function PicasaPhotos(ByVal s)
s=Replace(s,"https://","http://")
s=Replace(UTF8Decode(s),PicasaBase,"")
s=Replace(s,"&amp;","&")
s=rxReplace(s,"(kind|alt|hl)=(\w+)&?","")
q=rxSelect(s,"\?.*$"):s=rxReplace(s,"\?.*$","")
s=rxReplace(s,"^/|/$","")
s=rxReplace(s,"^(\w+)/(\d+)$","data/feed/base/user/$1/albumid/$2")
s=rxReplace(s,"^(\w+)/(\w+)$","data/feed/base/user/$1/album/$2")
s=IIf(InStr(1,s,"/")>0,Replace(s,"feed/api","feed/base"),"data/feed/base/user/"&s&"")
If q<>"" Then s=s&q
s=s&IIf(rxTest(s,"\?$")Or rxTest(s,"\&$"),"",IIf(rxTest(s,"\?"),"&","?"))
s=PicasaBase&s&"alt=rss&kind=photo&"
PicasaPhotos=s
End Function
Function GoogleSearchCount(ByVal sSearch)
Dim sResponse,oXML,output
output=""
sResponse=GoogleSearch_(sSearch,0)
Set oXML=LoadXML(sResponse)
If IsObject(oXML)Then
output=oXML.selectSingleNode("//estimatedTotalResultsCount").Text
End If
Set oXML=Nothing
sResponse=""
GoogleSearchCount=output
End Function
Function GoogleLinks(ByVal sUrl)
GoogleLinks=GoogleSearchCount("link:"&sUrl&"")
End Function
Function GoogleCache(ByVal sUrl)
GoogleCache=GoogleSearchCount("site:"&sUrl&"")
End Function
Function GoogleSearch_(ByVal sSearchString,ByVal iStart)
sSearchString=Trim(sSearchString&"")
iStart=VOZ(iStart)
Dim sSOAP:sSOAP=""
sSOAP=sSOAP&"<?xml version='1.0' encoding='UTF-8'?>"
sSOAP=sSOAP&"<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/1999/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/1999/XMLSchema"">"
sSOAP=sSOAP&"<SOAP-ENV:Body>"
sSOAP=sSOAP&"<ns1:doGoogleSearch xmlns:ns1=""urn:GoogleSearch"" SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">"
sSOAP=sSOAP&"<key xsi:type=""xsd:string"">00000000000000000000000000000000</key>"
sSOAP=sSOAP&"<q xsi:type=""xsd:string"">shrdlu winograd maclisp teletype</q>"
sSOAP=sSOAP&"<start xsi:type=""xsd:int"">0</start>"
sSOAP=sSOAP&"<maxResults xsi:type=""xsd:int"">10</maxResults>"
sSOAP=sSOAP&"<filter xsi:type=""xsd:boolean"">true</filter>"
sSOAP=sSOAP&"<restrict xsi:type=""xsd:string""></restrict>"
sSOAP=sSOAP&"<safeSearch xsi:type=""xsd:boolean"">false</safeSearch>"
sSOAP=sSOAP&"<lr xsi:type=""xsd:string""></lr>"
sSOAP=sSOAP&"<ie xsi:type=""xsd:string"">latin1</ie>"
sSOAP=sSOAP&"<oe xsi:type=""xsd:string"">latin1</oe>"
sSOAP=sSOAP&"</ns1:doGoogleSearch>"
sSOAP=sSOAP&"</SOAP-ENV:Body>"
sSOAP=sSOAP&"</SOAP-ENV:Envelope>"
Dim oXML
Set oXML=LoadXml(sSOAP)
oXML.selectSingleNode("//key").Text=GoogleAPI
oXML.selectSingleNode("//q").Text=sSearchString
oXML.selectSingleNode("//start").Text=iStart
Dim oHTTP:Set oHTTP=XMLHTTP()
oHTTP.open "post","http://api.google.com/search/beta2",False
oHTTP.setRequestHeader "Content-Type","text/xml"
oHTTP.setRequestHeader "SOAPAction","doGoogleSearch"
oHTTP.send oXML
GoogleSearch_=oHTTP.responseText
End Function
Function GoogleSearch(ByVal sSearchString,ByVal iStart)
Dim output:output=""
output=XmlFormat(GoogleSearch_(sSearchString,iStart),"/fwx/SOAP/google/Google.xsl")
GoogleSearch=output
End Function
Function GoogleSiteSearch(ByVal sSearchString,ByVal iStart)
Dim sSiteSrch:sSiteSrch="site:"&MyDomain&""
If Not InStr(1,sSearchString,sSiteSrch)>0 Then
sSearchString=sSearchString&" "&sSiteSrch
End If
GoogleSiteSearch=GoogleSearch(sSearchString,iStart)
End Function
Function GoogleSpelling(ByVal sSearchString)
On Error Resume Next
Dim output:output=""
sSearchString=Trim(sSearchString&"")
Dim oXML
Set oXML=LoadXML("<?xml version='1.0' encoding='UTF-8'?>"&"<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/1999/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/1999/XMLSchema"">"&"  <SOAP-ENV:Body>"&"    <ns1:doSpellingSuggestion xmlns:ns1=""urn:GoogleSearch"""&"        SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">"&"      <key xsi:type=""xsd:string"">00000000000000000000000000000000</key>"&"      <phrase xsi:type=""xsd:string"">britney speers</phrase>"&"    </ns1:doSpellingSuggestion>"&"  </SOAP-ENV:Body>"&"</SOAP-ENV:Envelope>"&"")
oXML.selectSingleNode("//key").Text=GoogleAPI
oXML.selectSingleNode("//phrase").Text=sSearchString
Dim oHTTP:Set oHTTP=XMLHTTP()
oHTTP.open "post","http://api.google.com/search/beta2",False
oHTTP.setRequestHeader "Content-Type","text/xml"
oHTTP.setRequestHeader "SOAPAction","doSpellingSuggestion"
oHTTP.send oXML
Set oXML=LoadXml(oHTTP.responseText)
output=oXML.selectSingleNode("//return").Text
GoogleSpelling=output
End Function
Function GooglePR(ByVal s)
GooglePR=GetResponse("/fwx/php/library/PageRank.php?"&s&"")
End Function
Const GoogleLanguage_Server="http://64.233.179.104/translate_c?"
Function GoogleTranslate(ByVal u,ByVal tl)
GoogleTranslate=GoogleTranslate2(u,tl,"")
End Function
Function GoogleTranslateURL(ByVal u,ByVal tl)
GoogleTranslateURL=GoogleTranslateURL2(u,tl,"")
End Function
Function GoogleTranslate2(ByVal u,ByVal tl,ByVal sl)
Dim o:o=""
o=""
GoogleTranslate2=o
End Function
Function GoogleTranslateURL2(ByVal u,ByVal tl,ByVal sl)
Dim o:o=""
o=GetHumanResponse(GoogleLanguage_Server&"sl="&sl&"&tl="&tl&"&u="&URLEncode(u)&"&")
GoogleTranslateURL2=o
End Function
Class Google_Data_API
Public Email
Public Password
Public Authorization
Public Function Feed(ByVal URL)
Dim gXRQ
Set gXRQ=CreateObject("Microsoft.XMLHTTP")
gXRQ.open "GET",URL,false
gXRQ.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
gXRQ.setRequestHeader "Authorization","GoogleLogin auth="&Me.Authorization
gXRQ.Send()
Set Feed=LoadXML(gXRQ.responseText)
End Function
Public Function Login(Service)
Login=False
Dim gXRQ,I,RT
Set gXRQ=CreateObject("Microsoft.XMLHTTP")
gXRQ.open "POST","https://www.google.com/accounts/ClientLogin",false
gXRQ.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
gXRQ.send "Email="&Me.Email&"&Passwd="&Me.Password&"&service="&Service&"&source=Fyneworks-CMS-1.0&"
If gXRQ.status<300 Then
RT=gXRQ.responseText
I=InStr(RT,"Auth=")
If I=0 Then
DIE "Google response has no AUTH data"
Exit Function
Else
RT=Mid(RT,I+5,Len(RT))
I=InStr(RT,Chr(10)):If I>0 Then RT=Left(RT,I-1)
Login=True
Me.Authorization=RT
End If
End If
Set gXRQ=Nothing
End Function
Public Function SpreadsheetKey(ByVal s)
k=rxSelect(s,"(worksheets?|full|key)[\=\/][^/\?]+")
If k="" Then k=rxSelect(s,"^[\w\-]+")
k=rxSelect(k,"[\w\-]+$")
SpreadsheetKey=k
End Function
Public Function Spreadsheets()
url="http://spreadsheets.google.com/feeds/spreadsheets/private/full"
Set Spreadsheets=Me.Feed(url)
End Function
Public Function Worksheets(ByVal spreadsheet)
url="http://spreadsheets.google.com/feeds/worksheets/"&Me.SpreadsheetKey(spreadsheet)&"/private/values"&rxSelect(spreadsheet,"\?.+$")
Set Worksheets=Me.Feed(url)
End Function
Public Function Rows(ByVal worksheet)
url="http://spreadsheets.google.com/feeds/list/"&Me.SpreadsheetKey(worksheet)&"/"&rxSelect(worksheet,"\w+$")&"/private/values"&rxSelect(worksheet,"\?.+$")
Set Rows=Me.Feed(url)
End Function
End Class
Const GoogleAnalyticsCode_Live="http://www.google-analytics.com/ga.js"
Const GoogleAnalyticsCode_File="/$$TEMP/ga.js"
Function GoogleAnalyticsCode_Update()
DBG "### GoogleAnalyticsCode_Update"
Dim dt,ga
dt=FWX.Setting("google.analytics.ga.cache")
DBG "### GoogleAnalyticsCode_Update: dt = "&dt
DBG "### GoogleAnalyticsCode_Update: Expires in "&DateAdd("h",1,dt)
DBG "### GoogleAnalyticsCode_Update: Refresh if "&DateAdd("h",1,dt)&" > "&Now()
If CDate(DateAdd("h",1,dt))>Now()Then
DBG "### GoogleAnalyticsCode_Update: Refresh if "&DateAdd("h",1,dt)&" > "&Now()&" - Yes, this is true (refresh not required)"
End If
If dt<>"" Then
If CDate(DateAdd("h",1,dt))>Now()Then
DBG "### GoogleAnalyticsCode_Update: Cache is fresh, quit now."
Exit Function
End If
End If
DBG "### GoogleAnalyticsCode_Update: Refresh cache..."
ga=GetHumanResponse(GoogleAnalyticsCode_Live)
DBG "### GoogleAnalyticsCode_Update: Code received = "&Len(ga)&" bytes"
If InStr(1,ga,"Request failed")>0 Then Exit Function
WriteToFile GoogleAnalyticsCode_File,ga
FWX.Setting("google.analytics.ga.cache")=Now()
DBG "### GoogleAnalyticsCode_Update: Done (as of "&FWX.Setting("google.analytics.ga.cache")&")"
End Function
Function GoogleAnalyticsCode()
Call GoogleAnalyticsCode_Update()
Dim ga
ga=ReadFile(GoogleAnalyticsCode_File)
If ga="" Then
Redirect GoogleAnalyticsCode_Live
End If
DIE ga
End Function
Function GA_TagLinks(ByVal s,ByRef col)
If IsObject(col)Then
Dim t:Set t=New QuickString:t.Divider="&"
Dim p
For Each p In Split("source,medium,term,content,campaign",",")
If col("utm_"&p)<>"" Then t.Add p
Next
If t<>"" Then
Dim a,m,l,x
Set a=rxMatches(s,"<a.*?href=[""'][^""']*[""']")
For Each m In a
x=m
l=rxReplace(m,"^<a.*?href=[""']|[""']$","")
t.Reset()
For Each p In Split("source,medium,term,content,campaign",",")
If InStr(1,l,"utm_"&p)>0 Then
Else
If col("utm_"&p)<>"" Then t.Add "utm_"&p&"="&URLEncode(col("utm_"&p))&""
End If
Next
If t<>"" Then
If InStr(1,l,"?")>0 Then
If Right(l,1)="&" Then
l=l&t&""
Else
l=l&"&"&t&""
End If
Else
l=l&"?"&t&"&"
End If
x=rxReplace(m,"(^<a.*?href=[""'])[^""']*([""']$)","$1"&l&"$2")
End If
s=Replace(s,m,x)
Next
End If
End If
GA_TagLinks=s
End Function
Class Google_Checkout_Cart
Public MerchantId
Public MerchantKey
Public Environment
Public Data
Public Property Get ApiServer()
Dim s
s="checkout.google.com"
If UCase(Environment)="SANDBOX" Then s="sandbox.google.com/checkout"
ApiServer="https://"&s&"/"
End Property
Public Property Get CheckoutURL()
CheckoutURL=ApiServer&"api/checkout/v2/merchantCheckout/Merchant/"&MerchantId
End Property
Public Property Get OrderURL()
OrderURL=ApiServer&"api/checkout/v2/request/Merchant/"&MerchantId
End Property
Public Function XML()
XML=Me.Data
End Function
Public Function GetResponse()
GetResponse=SendRequest(Me.XML,CheckoutURL)
End Function
Public Sub Checkout()
ProcessSynchronousResponse Me.GetResponse()
End Sub
Public Sub ProcessSynchronousResponse(ResponseXml)
Dim ResponseNode
Set ResponseNode=GetRootNode(ResponseXml)
Select Case ResponseNode.Tagname
Case "request-received"
DisplayResponse ResponseXml
Case "error"
DisplayResponse ResponseXml
Case "diagnosis"
DisplayResponse ResponseXml
Case "checkout-redirect"
ProcessCheckoutRedirect ResponseXml
Case Else
End Select
End Sub
Public Sub DisplayResponse(ResponseXml)
Response.Write "<pre>"
Response.Write Server.HTMLEncode(ResponseXml)
Response.Write "</pre>"
End Sub
Public Sub ProcessCheckoutRedirect(ResponseXml)
Dim ResponseNode,RedirectUrl
Set ResponseNode=GetRootNode(ResponseXml)
RedirectUrl=GetElementText(ResponseNode,"redirect-url")
Response.Redirect RedirectUrl
Set ResponseNode=Nothing
End Sub
Function SendRequest(XmlData,PostUrl)
Dim XmlHttp,BasicAuthentication,ResponseXml
Set XmlHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
XmlHttp.Open "POST",PostUrl,false
Const SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS=2
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS=13056
XmlHttp.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS,(XmlHttp.getOption(SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS)-SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
XmlHttp.SetRequestHeader "Authorization","Basic "&Base64Encode(MerchantId&":"&MerchantKey)
XmlHttp.SetRequestHeader "Content-Type","application/xml; charset=UTF-8"
XmlHttp.SetRequestHeader "Accept","application/xml; charset=UTF-8"
XmlHttp.Send XmlData
ResponseXml=XmlHttp.ResponseText
SendRequest=ResponseXml
Set XmlHttp=Nothing
End Function
End Class
Const HostIP_Server="http://api.hostip.info/"
Function HostIP(ByVal sIP)
Dim sURL,sGEO
If IsEmpty(sIP)And(Not IsLocal)Then
sIP=VISITOR_IP
End If
sGEO=MySession("HostIP:"&sIP)
If IsSet(sGEO)Then
Else
sURL=HostIP_Server&"?ip="&sIP
sGEO=GetResponse(sURL)
MySession("HostIP:"&sIP)=sGEO
End If
Dim oGEO
Set oGEO=LoadXml(sGEO)
Set HostIP=oGEO
End Function
Function HostIPNode(ByVal sIP,ByVal sN)
On Error Resume Next
Dim s,oGEO,oNOD
Set oGEO=HostIP(sIP)
If IsObject(oGEO)And Err.Number=0 Then
Set oNOD=oGEO.selectSingleNode("//"&sN&"")
If IsObject(oNOD)And Err.Number=0 Then
s=oNOD.Text
End If
End If
Err.Clear()
HostIPNode=s
End Function
Function HostIPCode(ByVal sIP)
HostIPCode=TOT(HostIPNode(sIP,"Hostip/countryAbbrev"),"UK")
End Function
Function HostIPCity(ByVal sIP)
HostIPCity=HostIPNode(sIP,"Hostip/gml:name")
End Function
Function HostIPCountry(ByVal sIP)
HostIPCountry=HostIPNode(sIP,"Hostip/countryName")
End Function
Function IPCoords(ByVal sIP)
Dim sCoords:sCoords=""
Dim dOutput:Set dOutput=Dictionary()
s=HostIPNode(sIP,"gml:coordinates")
If s<>"" Then
s=Split(s,",")
dOutput("long")=s(0)
dOutput("lat")=s(1)
End If
Set IPCoords=dOutput
End Function
Function IPLong(ByVal sIP)
Set oC=IPCoords(sIP)
IPLong=oC("long")
Set oC=Nothing
End Function
Function IPLat(ByVal sIP)
Set oC=IPCoords(sIP)
IPLat=oC("lat")
Set oC=Nothing
End Function
Const GeoIPLoc_Server="/fwx/php/geoiploc/lookup.php"
Function GeoIPLoc(ByVal sIP)
If IsEmpty(sIP)Then sIP=VISITOR_IP
If j("GeoIPLoc")="" Then
j("GeoIPLoc")=Request.ServerVariables("HTTP_CF-IPCOUNTRY")
End If
If j("GeoIPLoc")="" Then
j("GeoIPLoc")=MySession("GeoIPLoc:"&sIP)
End If
If j("GeoIPLoc")="" Then
j("GeoIPLoc")=GetResponse(GeoIPLoc_Server&"?ip="&sIP)
MySession("GeoIPLoc:"&sIP)=sGEO
End If
GeoIPLoc=j("GeoIPLoc")
End Function
Function GeoIPLocNode(ByVal sIP,ByVal sN)
Dim s,sGEO,aGEO,oGEO
sGEO=GeoIPLoc(sIP)
If InStr(1,sGEO,"<")>0 Then
s=""
Else
aGEO=Split(sGEO,",")
If UBound(aGEO)>0 Then
Set oGEO=Dictionary()
oGEO("code")=aGEO(0)
oGEO("abbr")=aGEO(1)
oGEO("name")=aGEO(2)
s=oGEO(sN)
Set oGEO=Nothing
End If
End If
GeoIPLocNode=s
End Function
Function GeoIPLocName(ByVal sIP)
GeoIPLocName=GeoIPLocNode(sIP,"name")
End Function
Function GeoIPLocCode(ByVal sIP)
Dim s
s=GeoIPLocNode(sIP,"code")
If s="GB" Then s="UK"
GeoIPLocCode=s
End Function
Function GeoIPLocCountry(ByVal sIP)
GeoIPLocCountry=GeoIPLocNode(sIP,"abbr")
End Function
Function ISBN(ByVal t)
End Function
Function xMethodExchangeRate(ByVal sCountry1,ByVal sCountry2)
Dim oXML:Set oXML=LoadXml("/fwx/SOAP/currency/request.xml")
oXML.selectSingleNode("//country1").Text=sCountry1
oXML.selectSingleNode("//country2").Text=sCountry2
Dim oHTTP:Set oHTTP=XMLHTTP()
oHTTP.open "post","http://services.xmethods.net:80/soap",False
oHTTP.setRequestHeader "Content-Type","application/xml"
oHTTP.setRequestHeader "SOAPAction","CurrencyExchangeService"
oHTTP.send oXML
Set oXML=Nothing
Dim output:output=oHTTP.responseText
Set oHTTP=Nothing
output=rxFind(output,"([0-9]+\.[0-9]+)\<\/Result\>")
output=rxFind(output,"([0-9]+\.[0-9]+)")
xMethodExchangeRate=output
End Function
Const YouTubeBase="http://www.youtube.com/"
Function YouTubeVideos(ByVal s)
s=Replace(s,"&amp;","&")
s=rxReplace(s,"^.*//","")
q=rxSelect(s,"\?.*$"):s=rxReplace(s,"\?.*$","")
s=rxReplace(s,"^/|/$","")
s=rxReplace(s,"^(www\.)?youtube\.com","gdata.youtube.com")
s=rxReplace(s,"/user/(\w+)$","/feeds/api/users/$1/uploads")
If q<>"" Then s=s&q
s=s&IIf(rxTest(s,"\?$")Or rxTest(s,"\&$"),"",IIf(rxTest(s,"\?"),"&","?"))
s="http://"&s
YouTubeVideos=s
End Function
CONST FB_API_Version="1.0"
CONST FB_API_RestServer="http://api.facebook.com/restserver.php"
CONST FB_CANVAS_PPrefix="fb_sig_"
Class FWX_Facebook_Rest_Client
Public ApiKey
Public SessionKey
Public SecretKey
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
End Sub
Public Function auth_createToken()
DBG "FB::auth_createToken..."
result=CallApiMethod("auth.createToken",Dictionary())
DBG "FB::auth_createToken="&Encode(result)
token=NodeValue(result,"auth_createToken_response")
If token="" Then
DIE "FB::auth_createToken failed"
End If
DBG "FB::auth_createToken-token="&token
auth_createToken=token
End Function
Public Function auth_getSession(auth_token)
DBG "FB::auth_getSession..."
Set out=Dictionary()
DBG "FB::auth_getSession-auth_token="&auth_token
result=CallApiMethod("facebook.auth.getSession","auth_token="&auth_token&"")
DBG "FB::auth_getSession-result="&Encode(result)
If InStr(result,"error_response")<>0 Then
DBG "FB::auth_getSession-failed: "&ErrorInfo(result)
Else
DBG "FB::auth_getSession-success: "&ErrorInfo(result)
Set xml=LoadXML(result)
SessionKey=NodeValue(xml,"session_key")
DBG "FB::auth_getSession-session_key: "&NodeValue(xml,"session_key")
out.add "session_key",NodeValue(xml,"session_key")
out.add "uid",NodeValue(xml,"uid")
expire=VOZ(NodeValue(xml,"expires"))
Set xml=Nothing
If expire=0 Then
out.add "expires","never"
Else
out.add "expires",UnixDate(expire)
End If
DBG "FB::auth_getSession-UID:"&out("uid")&" (exp:"&out("expires")&", "&UnixTime(out("expires"))&")="&out("session_key")&""
End If
Set auth_getSession=out
Set out=Nothing
End Function
Public Function CallApiMethod(strMethod,oParams)
Set oParams=QueryToDic(oParams)
oParams("method")=strMethod
data=GenerateRequestURI(oParams)
DBG "FB::CallApiMethod-<a target='_blank' href='"&FB_API_RestServer&"?"&data&"'>test</a>"
Dim oXMLHTTP
Set oXMLHTTP=Server.CreateObject("MSXML2.ServerXMLHTTP")
oXMLHTTP.Open "POST",FB_API_RestServer,False,"",""
oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
oXMLHTTP.Send data
DBG "FB::CallApiMethod-response:"&Encode(oXMLHTTP.ResponseText)
CallApiMethod=oXMLHTTP.ResponseText
Set oXMLHTTP=Nothing
End Function
Public Property Get RequestIsValid
Dim strItem,oRequestParams
Set oRequestParams=Server.CreateObject("Scripting.Dictionary")
For Each strItem In Request.Form
If(Left(strItem,Len(FB_CANVAS_PPrefix))=FB_CANVAS_PPrefix And Not strItem=FB_CANVAS_PPrefix)Then
oRequestParams(Mid(strItem,Len(FB_CANVAS_PPrefix)+1))=Request.Form(strItem)
End If
Next
RequestIsValid=(GenerateSig(oRequestParams)=Request.Form("fb_sig"))
Set oRequestParams=Nothing
End Property
Public Function ErrorInfo(xml)
ErrorInfo="Error "&NodeValue(xml,"error_code")&": "&NodeValue(xml,"error_msg")&""
End Function
Public Function CanvasRedirect(strURI)
CanvasRedirect="<fb:redirect url="""&strURI&""" />"
End Function
Public Function ErrorMessage(strMsg)
ErrorMessage="<fb:error message="""&strMsg&""" />"
End Function
Public Function SuccessMessage(strMsg)
SuccessMessage="<fb:success message="""&strMsg&""" />"
End Function
Private Function GenerateRequestURI(oParams)
If(Len(Application("FB_CallID"))=0)Then Application("FB_CallID")=100005
If oParams("session_key")="none" Then
oParams.remove "session_key"
Else
oParams("session_key")=SessionKey
End If
oParams("api_key")=ApiKey
oParams("call_id")=Application("FB_CallID")
oParams("v")="1.0"
Dim strItem
For Each strItem In oParams.Keys
GenerateRequestURI=GenerateRequestURI&strItem&"="&Server.UrlEncode(oParams(strItem))&"&"
Next
GenerateRequestURI=GenerateRequestURI&"sig="&GenerateSig(oParams)
Application("FB_CallID")=Application("FB_CallID")+205
End Function
Private Function GenerateSig(oParams)
Set oParams=SortDictionary(oParams)
Dim strSig,strItem
For Each strItem In oParams
strSig=strSig&strItem&"="&oParams(strItem)
Next
strSig=strSig&SecretKey
strSig=MD5(strSig)
GenerateSig=strSig
End Function
Public Function GenerateSig_cookies(cookies)
Set dict=Dictionary()
For Each item In cookies
dict.add item,cookies(item)
Next
GenerateSig_cookies=GenerateSig(dict)
Set dict=Nothing
End Function
End Class
Function PieChart(ByRef rD)
End Function
Function UseCaptcha()
Dim s,b
If IsLoggedIn Then
b=False
ElseIf IsStaffPC Then
b=False
Else
b=True
s=TOT(USE_CAPTCHA,FWX.Setting("captcha"))
If s<>"" Then b=IsTrue(s)
End If
UseCaptcha=b
End Function
Function reCAPTCHA_Write()
reCAPTCHA_Write="<style>.honeytrap{position:absolute!important;top:-9999px!important;left:-999px!important;}</style><input type='hidden' value='' name='honey' /><input type='text' value='' name='honey_email' class='honeytrap' /><input type='text' value='' name='honey_name' style='display:none;' /><p class='honeytrap'>Leave this empty: <input type='text' name='honey_url' /></p>"
End Function
Function reCAPTCHA_Check()
If(Not UseCaptcha)Then
j("recaptcha-check")="y"
Else
j("recaptcha-check")="n"
If Query.FQ("honey,honey_email,honey_name,honey_url")="" Then
j("recaptcha-check")="y"
End If
End If
reCAPTCHA_Check=IsTrue(j("recaptcha-check"))
End Function
Function PCDistance(ByVal a,ByVal b)
PCDistance=Geo_UK_PCDistance(a,b)
End Function
Function Geo_UK_PCDistance(ByVal p1,ByVal p2):Geo_UK_PCDistance=0
If Not IsObject(fwx)Then Exit Function
p1=DoPostcode(p1)
p2=DoPostcode(p2)
If p1="" Or p2="" Then Exit Function
Dim data:data=""
Dim objRequest
Set objRequest=HTTPRequest("/fwx/vb/geo/ukpc.asp?pc="&p1&"&to="&p2&"&")
objRequest.SetRequestHeader "FWX","http://"&SERVER_NAME&"/"
data=SendRequest(objRequest)
Set objRequest=Nothing
data=Split(data,",")
If UBound(data)>1 Then
Geo_UK_PCDistance=data(2)
Else
Geo_UK_PCDistance=0
End If
End Function
Const WikiWeb="http://en.wikipedia.org"
Function Wiki(ByVal s)
Dim o
s=WikiWeb&rxReplace("/wiki/"&s,"[\\/]+","/")
s=rxReplace(s,"/$","")
o=GetHumanResponse(s)
Wiki=o
End Function
Function WikiContent(ByVal s)
Dim o
o=Wiki(s)
o=rxSelect(o,"<!-- start content -->([\S\s]+)<!-- end content -->")
WikiContent=o
End Function
Function WikiCSS(ByVal s)
WikiCSS=WikiWeb&"/skins-1.5/monobook/main.css?52?jsview"
End Function
Function PingAll(ByVal url,ByVal title)
url=FullyQualifiedUrl(url)
j("Google")=PingGoogle(url)
j("GoogleSitemaps")=PingGoogleSitemaps(url&"gmap.xml")
j("Pingomatic")=Pingomatic(url,title)
j("Pingoat")=Pingoat(url,title)
j("Technorati")=PingTechnorati(url,title)
End Function
Function PingTechnorati(ByVal url,ByVal title)
On Error Resume Next
Dim str,svc:svc=""
url=FullyQualifiedUrl(url)
svc=svc&"http://technorati.com/ping/"
svc=svc&URLEncode(url)&""
str=GetResponse(svc)
PingTechnorati=Not CBool(InStr(1,str,"error")>0)
End Function
Function PingGoogle(ByVal url)
On Error Resume Next
Dim str,svc:svc=""
url=FullyQualifiedUrl(url)
svc=svc&"http://blogsearch.google.com/ping?"
svc=svc&"url="&URLEncode(url)&"&"
str=GetResponse(svc)
If Right(url,4)=".xml" Then
PingGoogleSitemaps url
End if
PingGoogle=Not CBool(InStr(1,str,"error")>0)
End Function
Function PingGoogleSitemaps(ByVal url)
On Error Resume Next
Dim str,svc:svc=""
url=FullyQualifiedUrl(url)
svc=svc&"http://www.google.com/webmasters/sitemaps/ping?"
svc=svc&"sitemap="&URLEncode(url)&"&"
str=GetResponse(svc)
PingGoogleSitemaps=Not CBool(InStr(1,str,"error")>0)
End Function
Function PingFeedBurner(ByVal url)
On Error Resume Next
Dim str,svc:svc=""
url=FullyQualifiedUrl(url)
svc=svc&"http://feedburner.google.com/fb/a/pingSubmit?"
svc=svc&"bloglink="&URLEncode(url)&"&"
str=GetResponse(svc)
PingFeedBurner=Not CBool(InStr(1,str,"error")>0)
End Function
Function Pingomatic(ByVal url,ByVal title)
On Error Resume Next
Dim str,svc:svc=""
url=FullyQualifiedUrl(url)
svc=svc&"http://pingomatic.com/ping/?"
svc=svc&"title="&URLEncode(title)&"&"
svc=svc&"blogurl="&URLEncode(url)&"&"
svc=svc&"rssurl="&URLEncode(url)&"rss.xml&"
svc=svc&"chk_weblogscom=on&"
svc=svc&"chk_blogs=on&"
svc=svc&"chk_technorati=on&"
svc=svc&"chk_feedburner=on&"
svc=svc&"chk_syndic8=on&"
svc=svc&"chk_newsgator=on&"
svc=svc&"chk_myyahoo=on&"
svc=svc&"chk_pubsubcom=on&"
svc=svc&"chk_blogdigger=on&"
svc=svc&"chk_blogrolling=on&"
svc=svc&"chk_blogstreet=on&"
svc=svc&"chk_moreover=on&"
svc=svc&"chk_weblogalot=on&"
svc=svc&"chk_icerocket=on&"
svc=svc&"chk_newsisfree=on&"
svc=svc&"chk_topicexchange=on"
svc=svc&"chk_google=on"
svc=svc&"chk_tailrank=on"
svc=svc&"chk_bloglines=on"
svc=svc&"chk_aiderss=on"
svc=svc&"chk_skygrid=on"
svc=svc&"chk_bitacoras=on"
str=GetResponse(svc)
Pingomatic=Not CBool(InStr(1,str,"error")>0)
End Function
Function Pingoat(ByVal url,ByVal title)
On Error Resume Next
Dim str,svc:svc=""
url=FullyQualifiedUrl(url)
svc=svc&"http://www.pingoat.com/index.php?pingoat=go&"
svc=svc&"blog_name="&URLEncode(title)&"&"
svc=svc&"blog_url="&URLEncode(url)&"&"
svc=svc&"rss_url="&URLEncode(url)&"rss.xml&"
svc=svc&"cat_0=1&id[]=0&id[]=1&id[]=2&id[]=3&id[]=4&id[]=5&id[]=6&id[]=7&id[]=8&id[]=9&id[]=10&id[]=11&id[]=12&id[]=13&id[]=14&id[]=15&id[]=16&id[]=17&id[]=18&id[]=19&id[]=20&id[]=21&id[]=22&id[]=23&id[]=24&id[]=25&cat_1=0&cat_2=0&"
str=GetResponse(svc)
Pingoat=Not CBool(InStr(1,str,"error")>0)
End Function
Class Dean_Edwards_Packer
Function Path()
Path=MyDomainCache&"packer/"
End Function
Function Script()
Script="/fwx/php/packer/pack.php"
End Function
Function Pack(ByVal s)
Dim dat
Dim out,key
key=DoEncrypt(MD5(s),"1234567890")
out=Me.Path()&key&".js"
BuildPath Me.Path()
DBG "PACKER.Pack: key="&key&", out="&out&""
ed=FWX.Setting("edition")
ca=FWX.Setting("PACKER.Cache")
If Left(ca,Len(ed))=ed Then
If InStr(1,ca,key&",")>0 Then
If LastModified(s)<LastModified(out)Then
dat=ReadFile(out)
End If
Else
ca=ca&key&","
End If
Else
ca=ed&","&key&","
End If
FWX.Setting("PACKER.Cache")=ca
DBG "PACKER.Pack: Cache="&ca&"<br/>dat="&Len(dat)&" bytes"
If Len(dat)>0 Then
If IsTest Then Response.AddHeader "X-PACKER",Len(dat)&" bytes from cache"
Else
Dim url,str
url=Me.Script()&"?src="&PhysicalPath(s)&"&out="&out&"&t="&TIMESTAMP&"&"
If IsTest Then Response.AddHeader "X-PACKER",url
str=GetResponse(url)
DBG "PACKER.Pack: response="&str
If 1=0 Then
ElseIf InStr(1,str,"PHP Warning")>0 Then
If IsTest Then Response.AddHeader "X-PACKER-ERROR","PHP Warning- "&out
ElseIf InStr(1,str,"PHP Error")>0 Then
If IsTest Then Response.AddHeader "X-PACKER-ERROR","PHP Error- "&out
ElseIf InStr(1,str,"Parse error")>0 Then
If IsTest Then Response.AddHeader "X-PACKER-ERROR","Parse error- "&out
ElseIf UCase(Left(str,2))="OK" Then
If IsTest Then Response.AddHeader "X-PACKER-OK","Cached in:"&out
dat=ReadFile(out)
End If
End If
DBG "PACKER.Pack: data="&Len(dat)&" bytes"
Pack=dat
End Function
Function PackCode(ByVal s)
PackCode=Code(s)
End Function
Function Code(ByVal s)
temp_file=Me.Path()&MD5(s)&".js"
WriteToFile temp_file,s
Code=Me.Pack(temp_file)
DeleteFile temp_file
End Function
End Class
Set JSObfuscator=New Dean_Edwards_Packer
Const FWX__Twillio_Base="https://api.twilio.com/2010-04-01"
Function Twillio_Call(caller,called,url)
accountSid=site.CFG("twillio.api")' "ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
authToken=site.CFG("twillio.token")' "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
baseUrl=FWX__Twillio_Base
smsUrl=baseUrl&"/Accounts/"&accountSid&"/Calls"
Set http=Server.CreateObject("MSXML2.ServerXMLHTTP")
http.open "POST",callUrl,False,accountSid,authToken
http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
sPostData="Caller="&Server.URLEncode(caller)
sPostData=sPostData&"&Called="&Server.URLEncode(called)
sPostData=sPostData&"&Url="&Server.URLEncode(url)
http.send sPostData
Twillio_Call=http.responseText
Set http=Nothing
End Function
Function Twillio_SMS(ByVal from,ByVal recipient,ByVal body)
Dim o,accountSid,authToken,baseUrl,smsUrl,http
accountSid=site.CFG("twillio.api")' "ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
authToken=site.CFG("twillio.token")' "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
baseUrl=FWX__Twillio_Base
smsUrl=baseUrl&"/Accounts/"&accountSid&"/SMS/Messages"
Set http=Server.CreateObject("MSXML2.ServerXMLHTTP")
http.open "POST",smsUrl,False,accountSid,authToken
http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
If from="" Or from="*" Or Len(from)<3 Then
from=site.CFG("twillio.number")
End If
body=TOE(body)
If Len(from)=0 Then
If IsTest Then
DBG.FatalError "SMS sender missing"
Else
DBG.Warning "SMS sender missing"
End If
Exit Function
End If
If Len(body)=0 Then
If IsTest Then
DBG.FatalError "SMS body empty, "&Len(body)&" chars"
Else
DBG.Warning "SMS body empty, "&Len(body)&" chars"
End If
Exit Function
End If
If Len(body)>160 Then
If IsTest Then
DBG.FatalError "SMS too long, "&Len(body)&" chars"
Else
DBG.Warning "SMS too long, "&Len(body)&" chars"
End If
Exit Function
End If
sPostData="From="&Server.URLEncode(from)
sPostData=sPostData&"&To="&Server.URLEncode(recipient)
sPostData=sPostData&"&Body="&Server.URLEncode(body)
http.send sPostData
o=http.responseText
Set http=Nothing
Set d=Dictionary()
d("notes")="To: "&recipient&vbCRLf&body&vbCRLf&""
d("log.quick")="y"
App.Actions.Log "SMS SENT",d
Set d=Nothing
Twillio_SMS=o
End Function
Function H404_SMS()
DIE "y"
End Function
Function H404__twillio_inbound()
UpdateCollection j,Query.Data
UpdateCollection j,Query.Fata
j("from")=TOT(j("From"),Query.FQ("From"))
j("to")=TOT(j("To"),Query.FQ("To"))
j("name")=TOT(j("CallerName"),Query.FQ("To"))
j("message")=j("Body")
j("reply")=j("message")
j("sid")=TOT(j("SmsSid"),Query.FQ("SmsMessageSid,MessageSid"))
j("url")="https://www.twilio.com/user/account/log/messages/"&j("sid")
If Not IsObject(app)Then Set app=New fwxAPP
If Not IsObject(shop)Then Set shop=New fwxSHOP
j("match")=j("from")
j("match")=rxReplace(j("match"),"^\+44","")
j("match")=rxReplace(j("match"),"^07","")
If j("match")<>"" Then
Set rA=shop.Accounts.Run("SELECT TOP 1 account,credentials,email,manager,supervisor,forename,name fullname FROM %TABLE WHERE REPLACE(t.[primary],'  ','') LIKE '%"&j("match")&"' OR REPLACE(t.[alternative],'  ','') LIKE '%"&j("match")&"' ORDER BY t.active DESC, t.date DESC")
If Exists(rA)Then
UpdateCollection j,rA
j("epos")=IIf(HasAccess(j("credentials"),"epos"),True,False)
If IsTrue(j("staff"))Then
j("from-label")="Staff (Stores) "
Else
j("staff")=IIf(HasAccess(j("credentials"),"staff"),True,False)
If IsTrue(j("staff"))Then
j("from-label")="Staff "
End If
End If
Else
Set rA=shop.Invoices.Jobs.Drivers.Run("SELECT TOP 1 %KEY,dbo.TOT(fullname,name) name,email FROM %TABLE WHERE REPLACE(t.[primary],'  ','') LIKE '%"&j("match")&"' OR REPLACE(t.[alternative],'  ','') LIKE '%"&j("match")&"' ORDER BY t.active DESC, t.date DESC")
If Exists(rA)Then
UpdateCollection j,rA
j("from-label")="Driver "'&j("name")
Else
Set rA=shop.Invoices.Jobs.Helpers.Run("SELECT TOP 1 %KEY,dbo.TOT(fullname,name) name,email FROM %TABLE WHERE REPLACE(t.[primary],'  ','') LIKE '%"&j("match")&"' OR REPLACE(t.[alternative],'  ','') LIKE '%"&j("match")&"' ORDER BY t.active DESC, t.date DESC")
If Exists(rA)Then
UpdateCollection j,rA
j("from-label")="Helper "'&j("name")
End If
End If
End If
KillRS rA
End If
If j("fullname")="" Then j("fullname")=j("CallerName")
j("notes")=j("message")&vbCRLf&vbCRLf&"From: "&j("from")&Enclose(j("fullname")," (A"&j("account")&" ",")")&"; "&vbCRLf&"To: "&j("to")&"; "&vbCRLf&""
j("log.quick")="y"
j("e")="twillio"
j("event")=IIf(j("type")="SMS","SMS RECEIVED","NUMBER CALLED")
j("name")="Twillio "&j("event")'IIf(j("type")="SMS","SMS RECEIVED","NUMBER CALLED")
j("pending")=1
App.Actions.Log j("name"),j
If j("notify")="" Then j("notify")=App.CFG("site.twillio.email")
If j("notify")="" Then j("notify")=App.CFG("site.email")
If j("notify")="" Then j("notify")=App.CFG("shop.email")
j("notify.subdomain")=TOT(App.CFG("site.twillio.subdomain"),"office")
If j("notify")<>"" Then
If GTZ(j("account"))Then
j("link")="http://"&rxReplace(SERVER_NAME,"^(dev|www)",j("notify.subdomain"))&"/x/accounts/"&j("account")
If IsStaff Or IsTrue(j("staff"))Then
j("link")="http://"&rxReplace(SERVER_NAME,"^(dev|www)",j("notify.subdomain"))&"/x/@shop.accounts.staff@/f?a=view&account="&j("account")
End If
ElseIf GTZ(j("driver"))Then
j("link")="http://"&rxReplace(SERVER_NAME,"^(dev|www)",j("notify.subdomain"))&"/x/drivers/?a=view&driver="&j("driver")
End If
SendEmail TOT(j("email"),"sms@"&MyDomain),j("notify"),"Twillio "&j("type")&": "&j("from-label")&Enclose(j("fullname"),""," ")&j("from"),Replace(j("notes"),vbCRLf,"<br/>")&"<br/><br/>"&Enclose(j("driver"),"Identified as: Driver ","<br/>")&Enclose(j("helper"),"Identified as: Helper ","<br/>")&Enclose(j("account"),"Identified as: Account ","<br/>")&Enclose(j("link"),"See: ","<br/>")&PrintColVars(j,"account,email,reply,sid,url")&"<br/>fwx.twillio."&j("type")&""
End If
End Function
Function H404_CALL_inbound()
j("type")="CALL"
Call H404__twillio_inbound()
RXML
DIE "<?xml version=""1.0""?><Response><Say>"&Enclose(j("forename"),"Hey ","! ")&"This is a special number for text messages only, but we've noted you called and somebody in our team will contact you as soon as possible.Thanks for calling!</Say><Sms>"&Enclose(j("forename"),"Hey ","! ")&"This number can't receive calls but we noted you called and will get back to you ASAP</Sms></Response>"
End Function
Function H404_SMS_inbound_failed()
j("dump")=Replace(Request.Form,"&",vbCRLf)
WriteToFile "/twillio/"&FmtDate(Now(),"%Y%M%D/%H%N%S_call_failed")&".txt",j("dump")
SendEmail "sms@"&MyDomain,"admin@"&MyDomain,"SMS "&Query.FQ("from"),j("dump")
RE
End Function
Function H404_SMS_inbound()
j("type")="SMS"
Call H404__twillio_inbound()
RXML
DIE "<?xml version=""1.0""?><Response><Sms>"&Enclose(j("forename"),"Hey ","! ")&"Thanks for your message, somebody in our team will read it and get back to you as soon as possible.</Sms></Response>"
End Function
Function H404_CALL_inbound_failed()
j("dump")=Replace(Request.Form,"&",vbCRLf)
WriteToFile "/twillio/"&FmtDate(Now(),"%Y%M%D/%H%N%S_call_failed")&".txt",j("dump")
SendEmail "sms@"&MyDomain,"admin@"&MyDomain,"CALL "&Query.FQ("from"),j("dump")
RE
End Function
CONST CC_API_Version="1.0"
CONST CC_API_RestServer="https://api.constantcontact.com/ws/customers/"
Class FWX_ConstantContact_Rest_Client
Public Username
Public Password
Public ApiKey
Public Secret
Public Function ApiServer()
ApiServer=CC_API_RestServer&Me.ApiKey&"%"&Me.Username&":"&Me.Password&""
End Function
Public Function AuthData()
AuthData="a="&Me.ApiKey&"&u="&Me.Username&"&p="&Me.Password&"&"
End Function
Public Function CallApiMethod(method,params)
DBG "CC::CallApiMethod"
DBG "CC::CallApiMethod- ApiServer="&Me.ApiServer
DBG "CC::CallApiMethod- AuthData="&Me.AuthData
DBG "CC::CallApiMethod- method="&method
DBG "CC::CallApiMethod- params="&params
Dim url,data
data=Me.AuthData&params
DBG "CC::CallApiMethod- data="&data
url=FullyQualifiedUrl("/fwx/php/ConstantContact/"&method&".php")
DBG "CC::CallApiMethod- url="&url
Dim oXMLHTTP,result
Set oXMLHTTP=Server.CreateObject("MSXML2.ServerXMLHTTP")
oXMLHTTP.Open "POST",url,False,"",""
oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
oXMLHTTP.Send data
DBG "CC::CallApiMethod-response:"&Len(oXMLHTTP.ResponseText)&" bytes"
result=oXMLHTTP.ResponseText
Set oXMLHTTP=Nothing
CallApiMethod=result
End Function
End Class
Const uv_w="UserVoice.working"
Function UservoiceXML(ByRef doc,ByVal field)
Dim o
o=XMLNodeText(doc,"ticket/"&field)
If o="" Then o=XMLNodeText(doc,field)
UservoiceXML=o
End Function
Function H404_UserVoice()
If Application.Contents(uv_w)=Minute(Now())Then
DIE "Busy"
End If
Application.Contents(uv_w)=Minute(Now())
DBG.Silence()
j("is_service")=1
j("bucket")="service"
j("e")="uservoice"
j("data")=Query.FQ("data")
j("event")=Query.FQ("event")
j("signature")=Query.FQ("signature")
j("log.quick")="y"
If j("data")<>"" Then
If Not IsObject(app)Then Set app=New fwxAPP
If Not IsObject(shop)Then Set shop=New fwxSHOP
If Not IsObject(site)Then Set site=New fwxSITE
Set doc=LoadXML(j("data"))
j("url")=UservoiceXML(doc,"url")
If j("ticket")="" Then j("ticket")=UservoiceXML(doc,"ticket_number")
If j("ticket")="" Then j("ticket")=UservoiceXML(doc,"id")
If j("subject")="" Then j("subject")=UservoiceXML(doc,"subject")
If j("subject")="" Then j("subject")=UservoiceXML(doc,"title")
If j("event")="new_ticket_note" Then
If j("notes")="" Then j("notes")=UservoiceXML(doc,"note/plaintext_body")
If j("notes")="" Then j("notes")=UservoiceXML(doc,"note/body")
j("email")=UservoiceXML(doc,"note/creator/email")
j("username")=UservoiceXML(doc,"note/creator/name")
Else
If j("notes")="" Then j("notes")=UservoiceXML(doc,"message/plaintext_body")
If j("notes")="" Then j("notes")=UservoiceXML(doc,"message/body")
End If
j("email")=UservoiceXML(doc,"contact/email")
j("is_admin_response")=UservoiceXML(doc,"is_admin_response")
If IsTrue(j("is_admin_response"))Then
If j("staff")="" Then j("staff")=UservoiceXML(doc,"sender/email")
If j("staff")="" Then j("staff")=UservoiceXML(doc,"creator/email")
If j("username")="" Then j("username")=UservoiceXML(doc,"sender/name")
Else
If j("username")="" Then j("username")=UservoiceXML(doc,"contact/name")
End If
If j("staff")=j("email")Then j("staff")=""
If j("postcode")="" Then j("postcode")=UservoiceXML(doc,"contact/location_region")
If j("postcode")="" Then j("postcode")=UservoiceXML(doc,"sender/traits/location_region")
If j("ua")="" Then j("ua")=UservoiceXML(doc,"contact/browser_name")
If j("ua")="" Then j("ua")=UservoiceXML(doc,"sender/traits/browser_name")
If j("os")="" Then j("os")=UservoiceXML(doc,"contact/os_name")
If j("os")="" Then j("os")=UservoiceXML(doc,"sender/traits/os_name")
If j("ua")="" Then j("ua")="Uservoice"
j("city")=UservoiceXML(doc,"contact/traits/location_city")
j("region")=UservoiceXML(doc,"contact/traits/location_region")
j("country")=UservoiceXML(doc,"contact/traits/location_country")
Set doc=Nothing
j("active")="y"
Select Case LCase(j("event"))
Case "new_ticket":j("event")="New Enquiry"
Case "new_ticket_reply":j("event")="Customer replied"
Case "new_ticket_admin_reply":j("event")="Staff replied"
Case "new_ticket_note":j("event")="Staff added note"
Case Else:
j("active")="n"
End Select
j("notes")=Trim(rxReplace(j("notes"),"[\t\n\r\s]+"," "))
j("notes")=rxReplace(j("notes"),"<[^>]+>","")
j("notes")=Slice(j("notes"),100)
j("notes")=Enclose(j("ticket"),"Ticket #",": ")&_
Enclose(rxReplace(j("event"),"^(Customer|Staff)",j("username")),"",": ")&j("notes")&""
If j("active")="y" Then
Dim rA
If j("staff")<>"" Then
j("staffAccount")=Application.Contents("staffAccount-"&j("staff"))
If j("staffAccount")="" Then
Set rA=shop.Accounts.Run("SELECT TOP 1 t.account staffAccount FROM %TABLE "&_
"WHERE t.email='"&SQLSafeString(j("staff"))&"' AND t.active=1 "&_
"ORDER BY t.editedon DESC "&_
"")
If Exists(rA)Then
UpdateCollection j,rA
End If
KillRS rA
Application.Contents("staffAccount-"&j("staff"))=j("staffAccount")
End If
End If
If j("email")<>"" Then
Set rA=shop.Accounts.Run("SELECT TOP 1 t.account, t.forename username, t.name fullname, i.invoice invoice, t.last_invoice_date, t.manager, t.manager accountSupervisor, i.supervisor invoiceSupervisor FROM %TABLE "&_
"LEFT JOIN tblSHOP_Invoice i ON i.invoice=t.last_invoice AND i.active=1 "&_
"WHERE t.email='"&SQLSafeString(j("email"))&"' AND t.active=1 "&_
"ORDER BY t.editedon DESC "&_
"")
If Exists(rA)Then
UpdateCollection j,rA
If IsDate(j("last_invoice_date"))Then
If CDate(j("last_invoice_date"))<Now()-180 Then
j("invoice")=""
j("invoiceSupervisor")=""
Set rQ=shop.Leads.Run(_
"SELECT TOP 1 t.quote, t.supervisor quoteSupervisor FROM %TABLE WHERE t.active=1 AND t.account="&VOZ(j("account"))&" AND datediff(d,t.date,getdate()) BETWEEN 0 AND 180 ")
If Exists(rQ)Then
UpdateCollection j,rQ
End If
KillRS rQ
End If
End If
If j("username")="" Then j("username")=j("fullname")
Else
KillRS rA
With shop.Accounts
.ID=0
If IsTrue(j("is_admin_response"))Then
If j("name")="" Then j("name")=Replace(rxSelect(j("email"),"^[^@]+"),"[\-\.\_\,]+"," ")
Else
If j("name")="" Then j("name")=j("username")
End If
If j("city")="Dallas" Then
j("city")=""
End If
j("crmstage")=VOV(Application.Contents("uservoice.default.crmstage"),j("uservoice.default.crmstage"))
j("ref")="uservoice"
.ID=.Save(j)
j("account")=.ID
End With
End If
End If
KillRS rA
End If
j("supervisor")=TOT(TOT(TOT(j("staffAccount"),j("invoiceSupervisor")),j("quoteSupervisor")),j("supervisor"))
If j("active")="y" Then
j("___n")=j("event")'Trim("UserVoice Ticket "&j("ticket")&Enclose(rxReplace(j("event"),"^(Customer|Staff)",j("username")),": ",""))
App.Actions.Log j("___n"),j
End If
Application.Contents(uv_w)=""
DIE "OK"
Else
Application.Contents(uv_w)=""
If IsTest Then
DIE "<form action='uservoice.asp' method='post'><input name='event' value='new_ticket'><textarea name='data'></textarea><button type='submit'>Test</button></form>"
End If
End If
End Function
Function BO_uservoice_inspector__START()
DBG.Silence()
UpdateCollection j,Query.Data
RW "<!DOCTYPE html>"&vbCRLf&"<html>"&vbCRLf&"  <head>"&vbCRLf&"    <title>Fyneworks Inspector Gadget</title>"&vbCRLf&"    <link href=""https://cdn.uservoice.com/packages/gadget.css"" media=""all"" rel=""stylesheet"" type=""text/css"" />"&vbCRLf&"  </head>"&vbCRLf&"  <body style='margin-right:10px'>"&vbCRLf&"      <table cellspacing='7' width='100%'>"&vbCRLf&""
If Not IsStaff Then
RW "<tr><td>ERROR</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/start'>You must be logged in to Fyneworks</a></td></tr>"
Else
j("sso-needed")=site.CFG("UserVoice.sso")
j("api-needed")=site.CFG("UserVoice.api")
If j("sso")<>j("sso-needed")Then
DBG.Warning "UserVoice Inspector Gadget Invalid SSO (needed "&j("sso-needed")&")"
RW "<tr><td>ERROR</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/config/site/UserVoice'>Invalid SSO Key</a></td></tr>"
ElseIf j("api")<>j("api-needed")Then
DBG.Warning "UserVoice Inspector Gadget Invalid API Key (needed "&j("api-needed")&")"
RW "<tr><td>ERROR</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/config/site/UserVoice'>Invalid API Key</a></td></tr>"
ElseIf j("fyn")<>site.CFG("UserVoice.fyn")Then
DBG.Warning "UserVoice Inspector Gadget Invalid FYN Key"
RW "<tr><td>ERROR</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/config/site/UserVoice'>Invalid FYN Key</a></td></tr>"
Else
Dim rA
If j("email")<>"" Then
Set rA=shop.Accounts.FindFields("TOP 1 t.account, t.date, t.name, t.invoices, t.balance, t.last_invoice, t.last_invoice_date","t.email='"&SQLSafeString(j("email"))&"' AND t.active=1")
If Exists(rA)Then
j("match")="y"
UpdateCollection j,rA
End If
End If
KillRS rA
If j("match")="y" Then
j("ordered")=False
If GTZ(j("invoices"))And IsDate(j("last_invoice_date"))Then
If Abs(DateDiff("d",Now(),j("last_invoice_date")))<180 Then
j("ordered")=True
End If
End If
RW "<tr><td colspan=""2""><a target='_blank' style='font-weight:bold;' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/A"&j("account")&"'>"&j("name")&"</a></td></tr>"
RW "<tr><td><a target='_blank' style='font-weight:bold;' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/A"&j("account")&"'>A"&j("account")&"</a></td><td align=""right"">"&DoDate(j("date"))&"</td></tr>"
If VOZ(j("balance"))<>0 Then
RW "<tr><td>Balance</td><td align=""right""><strong style='"&IIf(LTZ(j("balance")),"color:red;",IIf(GTZ(j("balance")),"color:green;","font-weight:normal;color:#999;"))&"'>"&VOZ(j("balance"))&"</strong></td></tr>"
End If
If GTZ(j("invoices"))Then
RW "<tr><td>Invoices</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/invoices?account="&j("account")&"'>"&j("invoices")&" invoices</a></td></tr>"
Set rX=shop.Invoices.FindFields("t.invoice,t.date","t.active=1 AND t.account="&VOZ(j("account")))
Do While Exists(rX)
If Abs(DateDiff("d",Now(),rX("date")))>180 Then
RW "<tr><td>INV"&rX("invoice")&"</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/invoices/"&rX("invoice")&"'>"&FmtDate(rX("date"),"%b '%y")&"</a></td></tr>"
Else
j("active-orders")=VOZ(j("active-orders"))+1
RW "<tr style=""color:#090;background:#f7f7f7""><td>INV"&rX("invoice")&"</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/invoices/"&rX("invoice")&"'>"&FmtDate(rX("date"),"%d %b")&"</a></td></tr>"
Set rJ=shop.Invoices.Jobs.FindFields("t.name,t.date,t.time,t.driverName,t.status,t.reference","t.active=1 AND t.invoice="&VOZ(rX("invoice")))
Do While Exists(rJ)
If rJ("date")>Now()And Abs(DateDiff("d",Now(),rJ("date")))>2 Then
RW "<tr style=""color:#090;background:#f7f7f7""><td>"&rJ("name")&"</td><td align=""right"">"&FmtDate(rJ("date"),"%d %b")&"</td></tr>"
Else
RW "<tr style=""color:#090;background:#f7f7f7""><td>"&rJ("name")&"</td><td align=""right"">"&FmtDate(rJ("date"),"%d %b")&"</td></tr>"
RW "<tr style=""color:#090;background:#f7f7f7""><td align=""right"" colspan=""2"">"&TOT(rJ("driverName"),"Driver not assigned yet")&"</td></td></tr>"
RW "<tr style=""color:#090;background:#f7f7f7""><td align=""right"" colspan=""2"">"&TOT(rJ("status"),"Unknown status")&"</td></td></tr>"
If IsSet(rJ("reference"))Then
RW "<tr style=""color:#090;background:#f7f7f7""><td align=""right"" colspan=""2"">"&TOT(rJ("reference"),"")&"</td></td></tr>"
End If
End If
rJ.MoveNext()
Loop
KillRS rJ
End If
rX.MoveNext()
Loop
KillRS rX
End If
If GTZ(j("active-orders"))Then
Else
Set rX=shop.Quotes.FindFields("TOP 5 t.quote,t.date","t.active=1 AND t.amount>0 AND t.account="&VOZ(j("account"))&" AND t.invoice=0 AND datediff(d,t.date,getdate())<=180")
If Exists(rX)Then
RW "<tr><td>Quotes</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/quotes?account="&j("account")&"'>"&rX.RecordCount&" latest quotes</a></td></tr>"
Do While Exists(rX)
If Abs(DateDiff("d",Now(),rX("date")))>180 Then
RW "<tr><td>INV"&rX("invoice")&"</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/invoices/"&rX("invoice")&"'>"&FmtDate(rX("date"),"%b '%y")&"</a></td></tr>"
Else
j("active-orders")=VOZ(j("active-orders"))+1
RW "<tr style=""color:#090;background:#f7f7f7""><td>Q"&rX("quote")&"</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/quotes/"&rX("quote")&"'>"&FmtDate(rX("date"),"%d %b")&"</a></td></tr>"
End If
rX.MoveNext()
Loop
End If
KillRS rX
End If
End If
End If
If j("email")<>"" Then
RW "<tr><td>Search</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/search?q="&URLENcode(j("name"))&"'>By name</a> | <a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&"/x/search?q="&URLENcode(j("email"))&"'>By email</a></td></tr>"
End If
If False Then
RW "<tr><td>...</td><td align=""right""><a target='_blank' href='"&TOT(PROTOCOL,"http")&"://"&rxReplace(SERVER_NAME,"^(dev|www)","office")&FULLREQUEST&"'>POPUP</a></td></tr>"
End If
End If
RW "      </table>"&vbCRLf&"      <script type=""text/javascript"">"&vbCRLf&"        //window.gadgetNoData = true;"&vbCRLf&"      </script>"&vbCRLf&"    <script src=""https://cdn.uservoice.com/packages/gadget.js"""&vbCRLf&"      type=""text/javascript""></script>"&vbCRLf&"  </body>"&vbCRLf&"</html>"&vbCRLf&""
j("link")=rxReplace(Query.Data("ref"),"^[\s\,]+","")
j("notes")="Ticket "&j("ticket[ticket_number]")&" - "&j("ticket[custom_fields][Nature of your enquiry]")'&"<br/>"&PrintCol(Query.Data)
If False Then
If IsStaff Then
If GTZ(j("invoice"))Or GTZ(j("account"))Then
App.Actions.Log "STAFF VIEWED TICKET IN USERVOICE",j
End If
End If
End If
RE
End Function
Dim FWX_Cronjob___DATA
Function H404_cron_exec():Call H404_cronjobs_exec():End Function
Function H404_cronjob_exec():Call H404_cronjobs_exec():End Function
Function H404_x_cronjob_exec():Call H404_cronjobs_exec():End Function
Function H404_x_cronjobs_exec():Call H404_cronjobs_exec():End Function
Function H404_cronjobs_exec()
Dim f
For Each f In Split("cronjob,cronkey,test,preview",",")
j(f)=Query.FQ(f)
Next
If VOZ(j("cronjob"))=0 Then
DBG.FatalError "Can't run cron job: missing job ID"
End If
If j("cronkey")="" And j("test")="" Then
DBG.FatalError "Can't run cron job: missing credentials"
End If
If Not IsObject(app)Then Set app=New fwxAPP
If Not IsObject(tmpl)Then Set tmpl=New fwxTMPL
If Not IsObject(site)Then Set site=New fwxSITE
If Not IsObject(quot)Then Set quot=New fwxQUOT
If Not IsObject(shop)Then Set shop=New fwxSHOP
Dim rJ
Set rJ=App.CronJobs.ByID(VOZ(j("cronjob")))
If Not Exists(rJ)Then
KillRS rJ
DBG.FatalError "Can't run cron job: job "&j("cronjob")&" not found!"
End If
Dim dJ
Set dJ=RS2DIC(rJ)
KillRS rJ
If j("cronkey")<>dJ("cronkey")And j("test")="" Then'CRONJOB____KEY(dJ)Then 'SHA256(dJ("date")&"*"&dJ("cronjob")&"*"&dJ("uuid"))Then
DBG.FatalError "Can't run cron job: invalid key"
End If
If IsSet(dJ("data"))Then
dJ("data-result")=parser.Parse(dJ("data"))
UpdateCollection Query.Fata,dJ("data-result")
Query.Fata("args")=""
Query.Fata("props")=""
Query.SetDates()
j("data")=CBool(dJ("data-result")<>"")
End If
If GTZ(dJ("action"))Then
Dim rA
Set rA=app.Actions.FieldsByID("t.name,t.action,t.staff,t.email,t.supervisor,t.supervisors,t.account,t.quote,t.invoice,t.url,t.notes",dJ("action"))
If Exists(rA)Then
UpdateCollection dJ,rA
Else
KillRS rA
DBG.FatalError "Cronjob failed: task "&dJ("action")&" could not be found"
End If
KillRS rA
For Each f In Split("quote,account,invoice,transaction,merchant,stockist",",")
If VOZ(dJ(f))=0 Then dJ(f)=""
Next
dJ("supervisors")=TOT(NList(dJ("supervisor")&", "&dJ("supervisors")),"0")
dJ("supervisors")=NList(dJ("supervisor")&", "&dJ("supervisors"))
If dJ("supervisors")<>"0" Then
Set rA=shop.accounts.FindFields("t.email","t.account IN ("&dJ("supervisors")&")")
If Exists(rA)Then
dJ("supervisorEmails")=rxReplace(rA.GetString(,,"",",",""),"[\r\n\s\t\,\;]+",";")
End If
KillRS rA
Set rA=shop.accounts.FindFields("t.[primary],t.[alternative]","t.account IN ("&dJ("supervisors")&") OR (t.email<>'' AND t.email IN ('"&SQLSafeString(dJ("staff"))&"'))")
If Exists(rA)Then
dJ("supervisorPhones")=rxReplace(rA.GetString(,,"",",",""),"[\r\n\s\t\,\;]+",";")
End If
KillRS rA
Else
If IsEmail(dJ("staff"))Then
Set rA=shop.accounts.FindFields("t.[primary],t.[alternative]","%SITEFTR AND t.email<>'' AND t.email IN ('"&SQLSafeString(dJ("staff"))&"')")
If Exists(rA)Then
dJ("supervisorPhones")=rxReplace(rA.GetString(,,"",",",""),"[\r\n\s\t\,\;]+",";")
KillRS rA
Set rA=shop.accounts.FindFields("t.[name]","%SITEFTR AND t.email<>'' AND t.email IN ('"&dJ("staff")&"')")
If Exists(rA)Then
dJ("supervisorName")=rA("name")
End If
KillRS rA
End If
KillRS rA
Else
End If
End If
dJ("remind-name")=dJ("name")&Enclose(dJ("account")," A","")&Enclose(dJ("quote")," Q","")&Enclose(dJ("invoice")," INV","")&Enclose(dJ("transaction")," T","")&" ("&dJ("action")&")"
j("notify.subdomain")=TOT(App.CFG("site.twillio.subdomain"),"office")
dJ("remind-link")="http://"&rxReplace(SERVER_NAME,"^(dev|www)",j("notify.subdomain"))&"/x/a/"&dJ("action")
dJ("remind-link-short")=ShortenURL(dJ("remind-link"))
dJ("email")=True
dJ("email_subject")=dJ("remind-name")
If dJ("supervisorName")<>"" Then
dJ("email_from")=dJ("supervisorName")&" <"&dJ("staff")&">"
Else
dJ("email_from")=dJ("staff")
End If
dJ("email_to")=dJ("staff")
dJ("email_cc")=dJ("supervisorEmails")
dJ("email_body")=dJ("remind-link")&" "&dJ("notes")
dJ("email_body")=TOT(MergeTemplate(BO__TMPL("!actions/reminder"),dJ),dJ("email_body"))
dJ("sms")=True
dJ("sms_to")=dJ("supervisorPhones")
dJ("sms_body")=dJ("remind-name")&" "&TOT(dJ("remind-link-short"),dJ("remind-link"))&" "&dJ("notes")
dJ("slack")=True
dJ("slack_channel")=dJ("slack_channel")
dJ("slack_body")=dJ("remind-name")&" "&TOT(dJ("remind-link-short"),dJ("remind-link"))&" "&dJ("notes")&Enclose(dJ("staff")," by ","")&Enclose(dJ("supervisorEmails")," for ","")&Enclose(dJ("account")," A","")&Enclose(dJ("quote")," Q","")&Enclose(dJ("invoice")," INV","")&""
j("action")=CBool(dJ("action-result")<>"")
End If
If IsTrue(dJ("plugin"))Then
Set FWX_Cronjob___DATA=dJ
dJ("plugin-result")=EXE("FWX_Cronjob__"&rxReplace(dJ("plugin_name"),"[^a-zA-Z0-9\_]+","")&"( FWX_Cronjob___DATA )")
UpdateCollection dJ,FWX_Cronjob___DATA
j("plugin")=CBool(dJ("plugin-result")<>"")
End If
If IsTrue(dJ("email"))Then
Dim dE,rE
Set dE=Dictionary()
For Each f In Split("template,from,to,bcc,cc,subject,body,format",",")
dE(f)=TOT(dJ(f),dJ("email_"&f))
Next
If GTZ(dE("template"))Then
Set rE=tmpl.Emails.FieldsByID("[from],/*[to],[bcc],[cc],*/[subject],html_header+html_content+html_footer [body]/*,[format]*/",dE("template"))
If Exists(rE)Then
For Each f In Split("template,from,to,bcc,cc,subject,body",",")
If dE(f)="" Then dE(f)=RSFieldGet(rE,f)
Next
End If
KillRS rE
End If
If dE("content")="" Then
dE("content")=dE("body")
End If
If dE("content")="" Then
dE("content")="<p style='font-size:130%;'><strong>"&dJ("name")&"</strong></p>"&Enclose(dJ("url"),"<p>URL: <strong>","</strong></p>")&Enclose(dJ("description"),"<p style='font-style: italic'>","</p>")&Enclose(dJ("notes"),"<p style='color:#999'>","</p>")&""
End If
dE("_key")="cronjobs"
dE("_id")="cronjob"&j("cronjob")
dE("footer")=Enclose(dE("footer"),"","<div class='small' style='font-size:small; border-top:solid 1px #ccc; color:#999;'>Server time: "&TOT(ConvertDate(Now(),j("country")),Now())&"</div>")
For Each f In Split("from,to,bcc,cc,subject,body,attachments",",")
If dE(f)<>"" Then dE(f)=parser.Parse(TOT(MergeTemplate(dE(f),dE),dE(f)))
Next
If IsEmail(dE("from"))Then
Set rA=shop.accounts.Run("SELECT TOP 1 t.[name],t.email FROM %TABLE WHERE %SITEFTR AND t.email<>'' AND t.email IN ('"&dE("from")&"') ORDER BY t.editedon DESC")
If Exists(rA)Then
If LCase(dE("from"))=LCase(RSFieldGet(rA,"email"))Then
dE("from_name")=rA("name")
Else
DBG.Warning "CRONJOB: Almost put the wrong name on the sender of an email"
End If
End If
KillRS rA
If dE("from_name")<>"" Then
dE("from")=dE("from_name")&" <"&dE("from")&">"
End If
End If
If rxTest(dE("to"),MyDomain)Then
dE("max-width")="100%"
End If
Set dE=tmpl.SendEmailObj(dE)
dJ("email-result")="<hr/>"&DIC2JSON(dE)&"<hr/>"
Set dE=Nothing
j("email")=CBool(dJ("email-result")<>"")
End If
If IsTrue(dJ("sms"))Then
For Each f In Split("to,body",",")
dJ("sms_"&f)=parser.Parse(TOT(MergeTemplate(dJ("sms_"&f),dJ),dJ("sms_"&f)))
Next
dJ("sms_body")=Left(dJ("sms_body"),160)
dJ("sms-result")=Encode(Twillio_SMS("",dJ("sms_to"),dJ("sms_body")))
j("sms")=CBool(dJ("sms-result")<>"")
End If
dJ("slack_channel")=FWX.Setting("slack.slackbot.channel")
dJ("slack_bot")=FWX.Setting("slack.slackbot.url")
If IsTrue(dJ("slack"))And IsSet(dJ("slack_bot"))Then
If Not rxTest(dJ("slack_bot"),"\&channel")Then
If dJ("slack_channel")="" Then dJ("slack_channel")="#customer-service"
dJ("slack_bot")=dJ("slack_bot")&"&channel="&UrlEncode(dJ("slack_channel"))
End If
If dJ("slack_body")="" Then dJ("slack_body")=dJ("sms_body")
If dJ("slack_body")="" Then dJ("slack_body")="Reminder for "&dJ("staff")&": "&dJ("remind-name")&", "&dJ("remind-link")
If dJ("slack_body")<>"" Then dJ("slack_body")=TOT(MergeTemplate(dJ("slack_body"),dJ),dJ("slack_body"))
If dJ("slack_body")<>"" Then
dJ("slack-result")=SendHTTPPost(dJ("slack_bot"),dJ("slack_body"))
End If
j("slack")=CBool(dJ("slack-result")<>"")
End If
If IsTrue(dJ("run_once"))Then
CRONJOB____REM(dJ("reference"))
app.Cronjobs.FieldByID("active",dJ("cronjob"))=0
End If
If False Then SendEmail "dev@"&MyDomain,"diego.alto@gmail.com","CronJob "&dJ("job")&": "&dJ("name"),"Request.Form="&Request.Form&"<br/>"&vbCRLf&"Query.FString="&Query.FString&"<br/>"&vbCRLf&"Query.FQ('a')="&Query.FQ("a")&"<br/>"&vbCRLf&"Query.FQ('b')="&Query.FQ("b")&"<br/>"&vbCRLf&"HEADERS >>> <br/>"&Replace(Request.serverVariables("ALL_HTTP"),vbCRLf,"<br/>")&"<br/>"&vbCRLf&"JOB CONFIGURATION >>> <br/>"&PrintCol(dJ)&"<br/>"&vbCRLf&""
j("status")="y"
DIEvars "status,email,sms,slack,data,plugin,action"
End Function
Class c_FWX
Public SettingsFolder
Public Property Get Setting(ByVal s)
Dim x
If x="" And Left(s,1)<>"#" Then x=EXE("FWX__"&rxReplace(rxReplace(rxReplace(s,"^app\W",""),"\W+","_"),"^_+",""))
If x="" Then x=Application.Contents(Me.DomainPrefix()&s)
If x="" Then x=Application.Contents(s)
Setting=x
End Property
Public Property Let Setting(ByVal s,ByVal v)
If v="" Then
Application.Contents.Remove(Me.DomainPrefix()&s)
Else
Application.Contents(Me.DomainPrefix()&s)=v
End If
End Property
Public Property Set Setting(ByVal s,ByVal v)
If v Is Nothing Then
Application.Contents.Remove(Me.DomainPrefix()&s)
Else
Set Application.Contents(Me.DomainPrefix()&s)=v
End If
End Property
Public Property Get Config(ByVal s)
Config=Me.Setting(s)
End Property
Public Property Let Config(ByVal s,ByVal v)
Me.Setting(s)=v
End Property
Public Property Set Config(ByVal s,ByVal v)
Set Me.Setting(s)=v
End Property
Function Run(ByVal s)
Dim rs
ConnectRS rs,Me.ConnectionString
OpenRS rs,s
Set Run=rs
End Function
Private Sub Class_Initialize()
SettingsFolder="/$$TEMP/FWX/"&Encrypt(Reverse("FYNEWORKS"))&""
EXE "FWX_INITIALIZE"
Me.SetKeys()
If Me.Setting("BOOT-TIME")<>"" Then
j("BOOT-AGE")=DateDiff("n",DON(Me.Setting("BOOT-TIME")),Now())
Else
Me.Reboot()
End If
If VOZ(Me.Setting("site"))=0 Then' Me.Setting("domain")="" Then
Me.Login()
End If
EXE "FWX_START"
End Sub
Function HasDatabase()
HasDatabase=Not Me.FSOOnly
End Function
Function FSOOnly()
FSOOnly=Not Len(FWX.ConnectionString)>0
End Function
Private lclDataSource
Public Function DataSource()
Dim o
If lclDataSource<>"" Then
o=lclDataSource
Else
o=Me.Setting("DataSource")
If IsLocal And Request.Cookies("LiveOnLocal")<>"y" Then
o=TOT(Me.Setting("DataSource-Local"),o)
Else
If IsTest And Request.Cookies("LiveOnTest")<>"y" Then
o=TOT(Me.Setting("DataSource-Test"),o)
End If
End If
End If
If IsTest Then
DBG ""
End If
DataSource=o
End Function
Private lclConnectionString
Public Function ConnectionString()
Dim o
If lclConnectionString<>"" Then
o=lclConnectionString
Else
Dim d
d=Me.DataSource
If d<>"" Then
o=FWXDB(d)
Else
o=Me.Setting("ConnectionString")
If IsLocal And MyCookies("LiveOnLocal")<>"y" Then
o=TOT(Me.Setting("ConnectionString-Local"),o)
Else
If IsTest And MyCookies("LiveOnTest")<>"y" Then
o=TOT(Me.Setting("ConnectionString-Test"),o)
End If
End If
End If
End If
ConnectionString=o
End Function
Public Sub SetKeys()
End Sub
Public Function EditionKey()
EditionKey="EDITION"&Enclose(Replace(MyDomainData,".","-"),".","")
End Function
Public Function EditionFile()
EditionFile="/$DOMAINCACHE/EDITION.txt"
End Function
Public Function GetEditionFromSharedRAM()
If IsObject(REDIS)Then
GetEditionFromSharedRAM=REDIS(Me.EditionKey)
End If
End Function
Public Function GetEditionFromDatabase()
If IsObject(app)Then
GetEditionFromDatabase=app(Me.EditionKey)
End If
End Function
Public Function GetEditionFromRAM()
GetEditionFromRAM=Application.Contents(Me.EditionKey)
End Function
Public Function GetEditionFromFSO()
Dim en,ef
ef=Me.EditionFile()
en=ReadFile(ef)
GetEditionFromFSO=en
End Function
Public Function GetEdition()
Dim en
If en="" Then en=Me.GetEditionFromSharedRAM()
If en="" Then en=Me.GetEditionFromDatabase()
If en="" Then en=Me.GetEditionFromRAM()
If en="" Then en=Me.GetEditionFromFSO()
GetEdition=en
End Function
Public Function SetEditionSharedRAM(ByVal en)
If IsObject(REDIS)Then
REDIS(Me.EditionKey)=en
End If
End Function
Public Function SetEditionDatabase(ByVal en)
If IsObject(REDIS)Then
REDIS(Me.EditionKey)=en
End If
End Function
Public Function SetEditionRAM(ByVal en)
Application.Contents(Me.EditionKey)=en
End Function
Public Function SetEditionFSO(ByVal en)
WriteToFile Me.EditionFile(),en
End Function
Public Function SetEdition(ByVal en)
Me.SetEditionSharedRAM en
Me.SetEditionDatabase en
Me.SetEditionRAM en
Me.SetEditionFSO en
SetEdition=en
End Function
Public Function NewEdition()
Dim en
en=FmtDate(Now(),"%y%M%D%H%N%S")
en=rxReplace(en,"[^\d\-]","")
Me.SetEdition en
NewEdition=en
End Function
Public Function KillEdition()
Me.SetEdition ""
DeleteFile Me.EditionFile
End Function
Public Function Edition()
Dim en,ef
en=Me.GetEdition()
If en="" Then
en=Me.NewEdition()
End If
Edition=en
End Function
Function Revision()
Revision=Me.Timestamp()
End Function
Function Timestamp()
Timestamp=Replace(Time(),":","")
End Function
Public Function Reboot()
Me.Setting("EDITION")=Me.Edition()
Me.Setting("REVISION")=Me.Revision()' Replace(Time(),":","")'Replace(Timer(),".","")
Me.Load("system")
EXE "FWX_SETTINGS"
If Me.Setting("ConnectionString")="" Or Me.Setting("ConnectionString-Local")="" Then
Me.Setting("ConnectionString")=FWXDB(Me.Setting("DataSource"))
Me.Setting("ConnectionString-Local")=FWXDB(Me.Setting("DataSource-Local"))
End If
If Me.Setting("ConnectionString")="" Or Me.Setting("ConnectionString-Local")="" Then
For Each s In Split("ConnectionString,ConectionString,Connection,Conection,DB,Database,Conn,Con",",")
If Me.Setting("ConnectionString")="" Then Me.Setting("ConnectionString")=Me.Setting(s)
If Me.Setting("ConnectionString-Local")="" Then Me.Setting("ConnectionString-Local")=Me.Setting(s&"-Local")
If s<>"ConnectionString" Then:Application.Contents.Remove(s):Application.Contents.Remove(s&"-Local"):End If
Next
End If
Me.SetKeys()
Me.Setting("BOOT-TIME")=Now()
End Function
Public Sub Load(ByVal level)
Select Case LCase(level)
Case "0","system","fwx":
UpdateCollection Application.Contents,Me.SystemSettings
Me.Load "server"
Case "1","server","s":
UpdateCollection Application.Contents,Me.ServerSettings&vbCRLf&Me.SettingsScript("/SETTINGS.asp")
Me.Load "instance"
Case "2","instance","i":
UpdateCollectionWithPrefix Application.Contents,Me.InstanceSettings&vbCRLf&Me.SettingsScript(SITE_FOLDER&"SETTINGS.asp"),Me.DomainPrefix()
Me.Load "domain"
Case "3","domain","d":
UpdateCollectionWithPrefix Application.Contents,Me.DomainSettings&vbCRLf&Me.SettingsScript("/"&Replace(MyDomain,".","-")&"/SETTINGS.asp"),Me.DomainPrefix()
End Select
Me.Setting("configuration")="":Application.Contents.Remove("configuration")'=""
Me.Setting("args")="":Application.Contents.Remove("args")'=""
Me.SetKeys()
End Sub
Function SettingsScript(ByVal s)
Dim o:o=""
s=rxReplace(s,"\:\d+","")
s=rxReplace(s,"^/+","/")
If FileExists(s)Then
o=Server.Execute(s)
End If
SettingsScript=o
End Function
Public Property Get SystemSettings()
SystemSettings=(ReadFile(Me.SystemFile))
End Property
Public Property Let SystemSettings(ByVal s)
Me.Load "system"
End Property
Public Property Get SystemFile()
SystemFile="/fwx/config.fwx"
End Property
Public Property Get ServerSettings()
ServerSettings=(ReadFile(Me.ServerFile))
End Property
Public Property Let ServerSettings(ByVal s)
WriteToFile Me.ServerFile,(s)
Me.Load "server"
End Property
Public Property Get ServerFile()
ServerFile=SettingsFolder&"/#1S"&SafePath("MYSERVER")&".fwx"
End Property
Public Property Get InstanceSettings()
If Me.Setting("FWX-DB")<>"" Then
Dim rL,sL
sL="EXEC dbo.[LOGIN] '"&MyDomain&"', '"&Me.Serial&"'"
ConnectRS rL,Me.Setting("FWX-DB")
rL.Open sL
If Exists(rL)Then
If rL("error")<>"" Then
If IsLocal Or IsTemp Then
DBG.FatalError "<a href='http://www.fyneworks.com/'>Fyneworks Content Management</a> - FWX ERROR: "&rL("error")&IIf(IsLocal,"<br/>"&sL,"")&""
Else
End If
End If
Dim s:Set s=New QuickString
For Each f In rL.Fields
s.Add vbCRLf&IIf(f.Name<>"configuration",f.Name&"=","")&f.Value
Next
If IsLocal Or IsTemp Then
Else
WriteToFile Me.InstanceFile,(s)
End If
Else
End If
KillRS rL
End If
InstanceSettings=(ReadFile(Me.InstanceFile))
End Property
Public Property Let InstanceSettings(ByVal s)
WriteToFile Me.InstanceFile,(s)
Me.Load "instance"
End Property
Public Property Get InstanceFile()
InstanceFile=SettingsFolder&"/#1C"&SafePath(DoEncrypt("INSTANCE",MyDomain))&".fwx"
End Property
Public Property Get DomainSettings()
DomainSettings=(ReadFile(Me.DomainFile))
End Property
Public Property Let DomainSettings(ByVal s)
WriteToFile Me.DomainFile,(s)
Me.Load "domain"
End Property
Public Property Get DomainFile()
DomainFile=SettingsFolder&"/#3D"&SafePath(MyHostname)&".fwx"
End Property
Function Login()
If Me.Setting("ConnectionString")<>"" Then
Else
Exit Function
End If
TimeTakenForConnectionTest=Timer()
If TestConnection(Me.ConnectionString)Then
Else
Exit Function
End If
Dim domain
If domain="" Then domain=Me.Setting("domain.data."&DOMAIN_NAME)
If domain="" Then domain=Me.Setting("domain.data")
If domain="" Then domain=DOMAIN_NAME_FOR_DATA
If domain="" Then domain=Me.Setting("domain")
If domain="" Then domain=MyDomain
Dim SiteID
If SiteID="" Then SiteID=VOZ(Me.Setting("SiteID."&DOMAIN_NAME))
If SiteID="" Then SiteID=VOZ(Me.Setting("SiteID."&domain))
If SiteID="" Then SiteID=Me.Setting("SiteID")
If GTZ(SiteID)And domain<>"" Then
Me.Setting("site")=SiteID
Me.Setting("domain.data")=domain
Me.SetKeys()
Exit Function
End If
Dim rI,sI
sI="EXEC dbo.[pcdINSTANCE_Check] "&MySite&", '"&domain&"'"
Set rI=Me.Run(sI)
If Exists(rI)Then
SaveXml rI,MyDomainCache&"#I/"&SafePath("INST")&".fwx"
Else
DBG.Warning "SITE LOGIN FAILED"
Set rI=OpenXML(MyDomainCache&"#I/"&SafePath("INST")&".fwx")
If Exists(rI)And IsLocal Then
fwxError "CheckInstance Error: "&MyDomain&"","Using backup copy of instance at '"&MyDomain&"'."&vbCRLf&""
DBG.FatalError "WARNING: using instance backup"
End If
End If
If Not Exists(rI)Then
fwxError "CheckInstance Fatal Error: "&MyDomain&"","Instance validation failed ("&MySite&", '"&MyDomain&"')."&vbCRLf&""
DBG.FatalError "Instance validation failed ("&MySite&", '"&MyDomain&"')"
Else
If rI("error")<>"" Then
DBG.FatalError "<a href='http://www.fyneworks.com/'>SEO Friendly CMS</a> ERROR: "&rI("error")&IIf(IsLocal,"<br/>"&sI,"")&""
End If
Me.Setting("site")=rI("site")
Me.Setting("domain.data")=rI("domain")
Me.Setting("INSTANCE")=rI("instance")
Me.SetKeys()
End If
KillRS rI
End Function
Private S_erial
Public Property Get Serial()
If S_erial="" Then
S_erial=RemoveSymbols(DoEncrypt(MyDomain,"/#1C"&SafePath("INSTANCE")&".fwx"))
End If
Serial=S_erial
End Property
Private D_omainKey
Public Property Get DomainKey()
If D_omainKey="" Then
D_omainKey=GetDomainKey(MyDomain)
End If
DomainKey=D_omainKey
End Property
Public Function GetDomainKey(ByVal s)
GetDomainKey=RemoveSymbols(DoEncrypt(Reverse(s),s))
End Function
Private D_omainPrefix
Public Property Get DomainPrefix()
If D_omainPrefix="" Then
D_omainPrefix=GetDomainPrefix(MyHostname)
End If
DomainPrefix=D_omainPrefix
End Property
Public Function GetDomainPrefix(ByVal s)
GetDomainPrefix="-"&Right(DoEncrypt(Left(RemoveSymbols(s),7),RemoveSymbols(Right(s,7))),4)&"-"
End Function
End Class
Dim FWX:Set FWX=New c_FWX
Application.Contents("DOMAIN:"&SERVER_NAME)=MyDomain
Function FWX_AddHeader(ByVal lbl,ByVal val)
On Error Resume Next
If DBG.Show Then
Else
Response.AddHeader lbl,val
End If
End Function
Function FWX_ContentType(ByVal val)
On Error Resume Next
If DBG.Show Then
DBG "FWX_ContentType('"&lbl&"')="&val&""
Else
Response.ContentType=val
End If
End Function
Sub FWX_ResponseStatus(ByVal val)
On Error Resume Next
If DBG.Show Then
DBG "FWX_ResponseStatus="&val&""
Else
Response.Status=val
End If
End Sub
Sub FWX_ResponseCharset(ByVal val)
On Error Resume Next
If DBG.Show Then
DBG "FWX_ResponseCharset="&val&""
Else
Response.Charset=val
End If
End Sub
Dim G_sFormVars
Function FormVars()
If Not IsObject(G_sFormVars)Then
Dim p
Set G_sFormVars=New QuickString
If IsObject(Query)Then
For Each p In Query.Fata
G_sFormVars.Add p&"&"
Next
Else
For Each p In Request.Form
G_sFormVars.Add p&"&"
Next
End If
End If
FormVars=G_sFormVars.GetString()
End Function
Dim G_sQueryVars
Function QueryVars()
If Not IsObject(G_sQueryVars)Then
Dim p
Set G_sQueryVars=New QuickString
If IsObject(Query)Then
For Each p In Query.Data
G_sQueryVars.Add p&"&"
Next
Else
G_sQueryVars=""
Exit Function
End If
End If
QueryVars=G_sQueryVars.GetString()
End Function
Function AllVars()
AllVars=FormVars&QueryVars
End Function
Dim G_dFormItems
Function FormItems()
If Not IsObject(G_dFormItems)Then Set G_dFormItems=Dictionary()
For Each p In Query.Fata
G_dFormItems(p)=Query.Fata(p)
Next
Set FormItems=G_dFormItems
End Function
Dim G_dQueryItems
Function QueryItems()
If Not IsObject(G_dQueryItems)Then Set G_dQueryItems=Dictionary()
For Each p In Query.Data
G_dQueryItems(p&"")=Query.Data(p)
Next
Set QueryItems=G_dQueryItems
End Function
Function DoReloadVars(ByVal sVars,ByVal sMode)
Dim output:Set output=New QuickString
Dim sLoaded
Dim bLive:bLive=CBool(InStr(1,sVars,"~_&")>0)
Dim aVars,curVar,varc,curVal,sFinalOutput
aVars=Split(sVars,"&")
For Each curVar In aVars
curVal=""
curVar=Trim(curVar&"")
varc=Left(curVar,1)
If(curVar<>"")And(varc<>"~")And(curVar<>"_t")And(curVar<>"_REF")And(curVar<>"_REFQ")Then
If((varc="_" OR varc="-")And bLive)Then
Else
If InStr(1,curVar,"=")>0 Then
curVal=Mid(curVar,InStr(1,curVar,"=")+1,Len(curVar))
curVar=Left(curVar,InStr(1,curVar,"=")-1)
Else
curVal=Query.FQ(curVar)
End If
If(Not InStr(1,sLoaded,","&curVar&",")>0)Then
If curVal<>"" Then
If sMode="fields" Then
output.Add "<input type=""hidden"" name="""&curVar&""" value="""&Encode(curVal)&""" />"&vbCRLf
Else
x=curVar&"="&Server.URLEncode(curVal)&"&"
x=Replace(x,"%2D","-")
output.Add x
End If
Else
If Query.FQ("z0q")<>"" Then
If sMode="fields" Then
output.Add "<input type=""hidden"" name="""&curVar&""" value=""""/>"&vbCRLf
Else
output.Add curVar&"=&"
End If
End If
End If
sLoaded=sLoaded&","&curVar&","
End If
End If
End If
Next
If sMode="url" Then
sFinalOutput=HTTP404_CleanQuery(output)
Else
sFinalOutput=(output)
End If
DoReloadVars=sFinalOutput
End Function
Function ReloadVars(sVars)
ReloadVars=DoReloadVars(sVars,"url")
End Function
Function ReloadFields(sVars)
ReloadFields=DoReloadVars(sVars,"fields")
End Function
Function ReAssignVars(ByVal sVars,ByVal sNewVars)
Dim output:Set output=New QuickString
Dim aVars,curVar,aNewVars,skipVar,newvar,curVal,oldNewVar
Dim bLive:bLive=CBool(InStr(1,sNewVars,"~_&")>0)
sNewVars=Replace(sNewVars,"~DATES","&~date&~start&~end"&"&~on&~before&~after&~from&~until&~date-search&")
sNewVars=Replace(sNewVars,"+DATES","&date&start&end"&"&~on&~before&~after&~from&~until&")
sNewVars=sNewVars&"&~date-search&"
sNewVars=sNewVars&"&~name-search&"
sVars=Replace(sVars,"-ajax","")
aVars=Split(sVars,"&")
For Each curVar In aVars
curVar=Trim(curVar&"")
If curVar<>"" Then
skipVar=False
If bLive Then
If Left(curVar,1)="_" Or Left(curVar,1)="-" Then
skipVar=True
End If
End If
If Not skipVar Then
aNewVars=Split(sNewVars,"&")
For Each newvar In aNewVars
curVal=""
newvar=Trim(newvar&"")
If newvar<>"" Then
If InStr(1,newvar,"=")>0 Then
curVal=Mid(newvar,InStr(1,newvar,"=")+1,Len(newvar))
oldNewVar=newvar
newvar=Left(newvar,InStr(1,newvar,"=")-1)
End If
If Left(newvar,1)="~" Then
If Mid(newvar,2)=curVar Then
r.Pattern="(\&|^)+(\"&newvar&")(\&|$)+"
sNewVars=r.Replace(sNewVars,"&")
skipVar=True
End If
End If
If LCase(curVar)=LCase(newvar)And(curVal<>"")And(Not skipVar)Then
r.Pattern="(\&|^)+(\~)*("&oldNewVar&")(\&|$)+"
sNewVars=r.Replace(sNewVars,"&")
output.Add curVar&"="&curVal&"&"
skipVar=True
End If
End If
Next
End If
If Not skipVar Then
output.Add curVar&"&"
End If
End If
Next
aNewVars=Split(sNewVars,"&")
For Each newvar In aNewVars
skipVar=False
If newvar<>"" Then
curVal=""
If InStr(1,newvar,"=")>0 Then
curVal=Mid(newvar,InStr(1,newvar,"=")+1,Len(newvar))
newvar=Left(newvar,InStr(1,newvar,"=")-1)
Else
End If
If Left(newvar,1)="~" Then
skipVar=True
End If
If Not skipVar Then
output.Add newvar&"="&curVal&"&"
Else
End If
End If
Next
sFinalOutput=HTTP404_CleanQuery(output)
ReAssignVars=sFinalOutput
Err.Clear()
End Function
Function ReAssign(ByVal sNewVars)
ReAssign=ReloadVars(ReAssignVars(AllVars(),sNewVars))
End Function
Function ReAssignQuery(ByVal sNewVars)
ReAssignQuery=ReloadVars(ReAssignVars(QueryVars(),sNewVars))
End Function
Function ReAssignForm(ByVal sNewVars)
ReAssignForm=ReloadVars(ReAssignVars(FormVars(),sNewVars))
End Function
Function ReAssignAll(ByVal sNewVars)
ReAssignAll=ReloadVars(ReAssignVars(AllVars(),sNewVars))
End Function
Function ReSubmit(ByVal sNewVars)
ReSubmit=RAWURL&"?"&ReAssign(sNewVars)
End Function
Function ReSubmitOnly(ByVal sVarsToLoad)
ReSubmitOnly=RAWURL&"?"&ReLoadVars(sVarsToLoad)
End Function
Function ReSubmitQuery(ByVal sNewVars)
ReSubmitQuery=RAWURL&"?"&ReAssignQuery(sNewVars)
End Function
Function ReSubmitVars(ByVal sOldVars,ByVal sNewVars)
ReSubmitVars=RAWURL&"?"&ReloadVars(ReAssignVars(sOldVars,sNewVars))
End Function
Function ReSubmitForm(ByVal sNewVars)
ReSubmitForm=ReloadFields(ReAssignVars(FormVars(),sNewVars))
End Function
Function ReSubmitFields(ByVal sNewVars)
ReSubmitFields=ReloadFields(ReAssignVars(AllVars(),sNewVars))
End Function
Dim G_sVARS
Function VARS()
If G_sVARS="" Then G_sVARS=ReloadVars(AllVars())
VARS=G_sVARS
End Function
Dim G_sFULLREQUEST
Function FULLREQUEST()
If G_sFULLREQUEST="" Then
Dim sFR_URL,sFR_QRY
sFR_URL=RAWURL
sFR_QRY=VARS
If sFR_QRY<>"" Then
sFR_QRY="?"&sFR_QRY
End If
G_sFULLREQUEST=sFR_URL&sFR_QRY
End If
FULLREQUEST=G_sFULLREQUEST
End Function
Function THIS()
THIS=Server.URLEncode(FULLREQUEST)
End Function
Function HTTP404_CleanQuery(ByVal output)
output=output
Dim i404,iQuery
i404=InStr(1,output,"404;")
If i404>0 Then
iQuery=InStr(i404+1,output,"?")
If Not iQuery>0 Then iQuery=InStr(i404+1,output,"&")
If iQuery>0 Then
output=Left(output,i404-1)&Mid(output,iQuery+1)
End If
End If
HTTP404_CleanQuery=output
End Function
Sub IsClientConnected():CheckClientConnected():End Sub
Sub CheckClientIsConnected():CheckClientConnected():End Sub
Sub CheckClientConnected()
If Not(Response.IsClientConnected)Then
FWX_ResponseStatus 200
Response.Write "CLIENT DISCONNECTED"
Response.End()
End If
End Sub
Class fwxMySession
Public Prefix
Public Sub Clear()
If IsObject(Session)Then
For Each c In Session.Contents
Session.Contents(c)=""
Next
Session.Abandon()
Else
End If
End Sub
Public Property Let Item(ByVal sC,ByVal sV)
If IsObject(Session)Then
If sV<>"" Then
Session.Contents(sC)=sV
Else
Session.Contents.Remove(sC)
End If
End If
End Property
Public Default Property Get Item(ByVal sC)
If IsObject(Session)Then
Item=Session.Contents(Me.Prefix&sC)
End If
End Property
End Class
Dim MySession:Set MySession=New fwxMySession
Class fwxCookies
Public Prefix
Public ServerKey
Public ClientKey
Public BlockKey
Public Sub Class_Initialize()
Prefix=""
ServerKey=Reverse(RemoveSymbols(SITE_PATH))
ClientKey=Reverse(RemoveSymbols(VISITOR_IP))
BlockKey=Reverse(TOT(rxSelect(VISITOR_IP,"^([0-9]+)(\.[0-9]+)?"),"192.168"))
End Sub
Public Sub Clear()
For Each c In Request.Cookies
Response.Cookies(c)=""
Response.Cookies(c).Expires=Now()-1
Next
End Sub
Public Property Let Item(ByVal sC,ByVal sV)
Me.Scoped(Me.Prefix,sC)=sV
End Property
Public Default Property Get Item(ByVal sC)
Item=Me.Scoped(Me.Prefix,sC)
End Property
Public Property Let Cookie(ByVal sC,ByVal sV)
Me.Item(sC)=sV
End Property
Public Property Get Cookie(ByVal sC)
Cookie=Me.Item(sC)
End Property
Public Property Let Global(ByVal sC,ByVal sV)
Me.Scoped("",sC)=sV
End Property
Public Property Get Global(ByVal sC)
Global=Me.Scoped("",sC)
End Property
Public Property Let Scoped(ByVal sP,ByVal sC,ByVal sV)
If IsCDN Then Exit Property
If sV<>"" Then
If DBG.Show Then
DBG "Cookies("&sC&") ABORTED - can't set headers in debug mode"
Exit Property
End If
Response.Cookies(sP&sC)=sV
Response.Cookies(sP&sC).Path="/"
Response.Cookies(sP&sC).Expires=DateAdd("d",7,Now())
Else
Response.Cookies(sP&sC)=""
End If
End Property
Public Property Get Scoped(ByVal sP,ByVal sC)
If IsCDN Then Exit Property
Scoped=Request.Cookies(sP&sC)
End Property
Public Property Let Secure(ByVal sC,ByVal sV)
Me.Item(sC)=DoEncrypt(Me.ServerKey,sV)
End Property
Public Property Get Secure(ByVal sC)
Secure=DoDecrypt(Me.ServerKey,Me.Item(sC))
End Property
Public Property Let IPBased(ByVal sC,ByVal sV)
Me.Item(sC)=DoEncrypt(Me.ClientKey,sV)
End Property
Public Property Get IPBased(ByVal sC)
IPBased=DoDecrypt(Me.ClientKey,Me.Item(sC))
End Property
Public Property Let BlockBased(ByVal sC,ByVal sV)
Me.Item(sC)=DoEncrypt(Me.BlockKey,sV)
End Property
Public Property Get BlockBased(ByVal sC)
BlockBased=DoDecrypt(Me.BlockKey,Me.Item(sC))
End Property
Private lclKey
Public Property Get Key()
If lclKey="" Then
lclKey=Me.Item("k")&""
If lclKey="" Then
If lclKey="" Then lclKey=RemoveSymbols(Timer())
Me.Item("k")=lclKey
End If
End If
Key=lclKey
End Property
End Class
Dim MyCookies
Set MyCookies=New fwxCookies
MyCookies.Prefix=IIf(IsBack,"B","F")
Const ENV_Resource_Images_Resolve__NotFound="!"
Const ENV_Resource_Images_Resolve__NotVirtual="*"
Function ENV_Resource_Image(ByVal path)
Dim url:url=path
Dim rxResize:rxResize="/([0-9]+)x([0-9]+)"
Dim rxSeoTxt:rxSeoTxt="\(.+\-\)|\(\-.+\)|/fwx-[^\.]+|/--[^\.]+"
Dim rxIPaths:rxIPaths="^/"&"(fwx|x|U|i|\@|img|images?|static)/"&"((img|images?)/)?"
Dim found,folder,sResize,sSeoText,output
If rxTest(url,"^(https?\:)")Then
output=url
ElseIf rxTest(url,"^(data\:)")Then
output=url
Else
url=Replace(url,"/hotlink-ok","")
If rxTest(url,rxSeoTxt)Then
sSeoText=rxSelect(url,rxSeoTxt)
url=rxReplace(url,rxSeoTxt,"")
End If
If rxTest(url,rxResize)Then
ENV_Resource_Image=path
Exit Function
sResize=rxSelect(url,rxResize)
url=rxReplace(url,rxResize,"")
End If
If FileExists(url)Then
found=url
ElseIf rxTest(url,rxIPaths)Then
url=rxReplace(url,rxIPaths,"")
url=rxReplace(url,"\(.*\-\)","")
url=rxReplace(url,"\(\-.*\)","")
Dim domain_folders,system_folders
domain_folders="/U/ /static/ "&Replace(Replace(FE__TMPL_Folders,"/fe",""),";"," ")&" "&TMPL_FOLDER&" /"
system_folders="/fwx/"' /fwx/inc/"
If False Then
ElseIf rxTest(url,"^domain/")Then
url=rxReplace(url,"^domain/","")
folders=domain_folders
ElseIf rxTest(url,"^system/")Then
url=rxReplace(url,"^system/","")
folders=system_folders
Else
folders=domain_folders&" "&system_folders
End If
folders=Trim(folders)
For Each folder In Split(folders," ")
For Each subfolder In Split("images/"," ")
If found="" Then
If FileExists(folder&subfolder&url)Then
found=folder&subfolder&url
End If
End If
Next
If found="" Then
If FileExists(folder&url)Then
found=folder&url
End If
End If
Next
If found="" Then
folder="/fwx/images/menu/"
If FileExists(folder&url)Then
found=folder&url
End If
End If
found=TranslatedPath(found)
End If
If rxTest(found,"gif$")Then
Dim png
png=rxReplace(found,"gif$","png")
If FileExists(png)Then
found=png
End If
End If
If found<>"" Then
output=sResize&found
End If
End If
ENV_Resource_Image=output
End Function
Function ENV_Resource_Images_MapPath(ByVal content)
Dim s
s=content
s=ENV_Resource_Images_Resolve(s)
Dim imS
Dim ulA,imA
Dim ulB,imB
Set imS=rxMatches(s,"<img[^>]*?(src)=[""'](?!https?://)[^""']*?\.\w{3}[^""']*?(?=[""'])[^>]*?>")
For Each imA In imS
ulB="":imB=""
ulA=rxSelect(imA,"(src)=[""'](?!https?://)[^""']*?\.\w{3}[^""']*?(?=[""'])")
ulA=Mid(ulA,6,Len(ulA)-2)
If True Then
ulA=rxReplace(ulA,"/([0-9]+)x([0-9]+)","")
ulB=PhysicalPath(ulA)
Else
End If
If ulB<>"" Then
imB=Replace(imA,ulA,ulB)
End If
s=Replace(s,imA,imB)
Next
Set imS=Nothing
ENV_Resource_Images_MapPath=s
End Function
Function ENV_Resource_Images_Resolve(ByVal content)
Dim s
s=content
Dim d:For Each d In Split("DOMAIN_NAME,SERVER_NAME,MyDomain",",")
s=rxReplace(s,"(href|src)\=(""|')http://((www|local|temp|test|dev|office|admin|d\d|svn|wap|m|w\d)\.)*"&rxEscape(Eval(d))&"[^/]*/","$1=$2/")
Next
s=rxReplace(s,"(href|src)\=(""|')http://"&"localhost"&"[^/]*/","$1=$2/")
Dim bs,bx,bp
bs=Application.Contents("blob.store")
bx=Application.Contents("blob.paths")
Dim imS
Dim ulA,imA
Dim ulB,imB
Set imS=rxMatches(s,"<img[^>]*?(src)=[""'](?!https?://)[^""']*?\.\w{3}[^""']*?(?=[""'])[^>]*?>")
For Each imA In imS
ulB="":imB=""
ulA=rxSelect(imA,"(src)=[""'](?!https?://)[^""']*?\.\w{3}[^""']*?(?=[""'])")
If ulA="" Then
Else
ulA=Mid(ulA,6,Len(ulA)-2)
If InStr(1,ulA,"/hotlink-ok/")>0 And Query.EXT<>"pdf" Then
Else
ulB=Application.Contents("IMG:"&ulA)
If ulB="" Then
ulB=ENV_Resource_Image(ulA)
If ulB="" Then
Application.Contents("IMG:"&ulA)=ENV_Resource_Images_Resolve__NotFound
ulB=ENV_Resource_Images_Resolve__NotFound
ElseIf ulA=ulB Then
Application.Contents("IMG:"&ulA)=ENV_Resource_Images_Resolve__NotVirtual
Else
Application.Contents("IMG:"&ulA)=ulB
Application.Contents("IMG:"&ulB)=ENV_Resource_Images_Resolve__NotVirtual
End If
End If
If ulB="" Then
ElseIf ulB=ENV_Resource_Images_Resolve__NotFound Then
ElseIf ulB=ENV_Resource_Images_Resolve__NotVirtual Then
ElseIf ulA=ulB Then
Else
If bs<>"" Then
For Each bp In Split(TOT(bx,"/"),",")
If bp<>"" Then
If StartsWith(ulB,bp)Then
ulB=bs&ulB
End If
End If
Next
End If
imB=Replace(imA,ulA,ulB)
s=Replace(s,imA,imB)
End If
End If
End If
Next
Set imS=Nothing
s=Replace(s,"http://localhost/","http://"&SERVER_NAME&"/")
ENV_Resource_Images_Resolve=s
End Function
Class fwxRequest
Public SPECIAL
Public URL
Public Property Get FULLURL()
With Me
FULLURL=.RAWURL&IIf(.QUERY<>"","?"&.QUERY&"","")
End With
End Property
Public RAWURL
Public RAWQRY
Public RAWURI
Public URI
Public EXT
Public PAGE,PAGEID
Public PATH,PATHID
Public FULLID
Public Data
Public Fata
Public QString
Public FString
Private Sub Class_Initialize()
Set Data=Dictionary()
Set Fata=Dictionary()
End Sub
Public Function Initialize()
Me.LoadDefaults()
Dim sQuery:sQuery=Me.QueryString()
If IsAjax Then sQuery=sQuery&IIf(Right(sQuery,1)="&","","&")&"-ajax="&RemoveSymbols(Timer())&""
Me.RawURL=QUERY_URL
Me.RawQRY=QUERY_STR
Me.RawURI=QUERY_URI
Me.URI=QUERY_URL&Enclose(QUERY_STR,"?","")
Dim sUrl,iQ,iSQ
sUrl=QUERY_URL
t="\/(index|default)\.(aspx?|html?|php\d?)$"
If rxTest(sUrl,t)Then
If sUrl<>Request.ServerVariables("SCRIPT_NAME")Then
PermanentRedirect rxReplace(sUrl,t,"/")&IIf(sQuery<>"","?"&sQuery,"")
RE
Else
sUrl=rxReplace(sUrl,t,"/")
End If
End If
If Len(sUrl)>4 Then
EXT=Replace(rxFind(sUrl,"\.\w{2,4}$"),".","")
End If
If IsFront And Is404 Then
If rxTest(sUrl,"/+$")Then
If IsAjax Or IsPost Then
Else
PermanentRedirect rxReplace(sUrl,"/+$","")&Enclose(sQuery,"?","")
End If
End If
End If
If DBG.Show Then
Else
If EXT="" Then
Response.Charset=TOT(FWX_Charset,"UTF-8")
Response.ContentType="text/html"
Else
sUrl=rxReplace(sUrl,"\.(\w{2,4})$","/.$1")
If rxTest(EXT,"^(gif|png|jpg?e?|bmp|tiff?)$")Then
Response.ContentType="image/"&rxReplace(EXT,"jpg?e?","jpeg")
Else
Response.Charset=TOT(FWX_Charset,"UTF-8")
Select Case EXT
Case "xml":Response.ContentType="application/xml"
Case "csv":Response.ContentType="text/csv"
Case "xls":Response.ContentType="application/vnd.ms-excel"';charset=ISO-8859-1"
Case "doc":Response.ContentType="application/vnd.ms-word"';charset=ISO-8859-1"
Case "css":Response.ContentType="text/css"
Case "js","json":Response.ContentType="text/javascript"
Case "txt","tsv":Response.ContentType="text/plain"
End Select
End If
End If
End If
If Is404 Then
If rxTest(sUrl,"//+")Then
sUrl=rxReplace(sUrl,"//+","/")
If IsAjax Or(IsPost And IsHuman)Then
Else
FWX_AddHeader "No-Double-Slash",sUrl&IIf(sQuery<>"","?"&sQuery,"")
PermanentRedirect sUrl&Enclose(sQuery,"?","")
RE
End If
End If
If rxTest(sUrl,"@[\w\.]+@")Then
Dim sEnt
sEnt=rxFind(sUrl,"@[\w\.]+@")
sEnt=rxFind(sEnt,"[\w\.]+")
j("-ENT")=sEnt
sUrl=rxReplace(sUrl,"/?"&"@[\w\.]+@"&"/?","/")
End If
If rxTest(sUrl,"\(.+\)")Then
Dim sBin
sBin=rxFind(sUrl,"\(.+\)")
sBin=Mid(sBin,2,Len(sBin)-2)
j("bin")=sBin
sUrl=rxReplace(sUrl,"/?\(.+\)/?","/")
End If
If InStr(1,sUrl,"/archive/")>0 Then
Dim mapping,mappings,map,list
mappings=Array("year 2\d{3}","month "&rxMonthName&"","day \d{2}","weekday "&rxDayName&"")
For Each mapping In mappings
map=Split(mapping," ")
map(1)="/archive/("&map(1)&")/"
map(1)=rxSelect(sUrl,map(1))
If map(1)<>"" Then
sUrl=Replace(sUrl,map(1),"/archive/")
map(1)=Replace(LCase(rxSelect(map(1),"[^/]*/$")),"/","")
If rxTest(map(1),"\D+")Then
For Each list In Array(fwxMonthNames,fwxMonthAbbrev,fwxDayNames,fwxDayAbbrev)
list=LCase("|"&list&"|")
i=InStr(1,list,"|"&map(1)&"|")
If i>0 Then
list=Left(list,i)
map(1)=Len(rxReplace(list,"[^\|]",""))
End If
Next
End If
sQuery=map(0)&"="&map(1)&"&"&sQuery
End If
Next
sQuery="archive=y&"&sQuery
sUrl=Replace(sUrl,"/archive/","/")
End If
Dim sPaging
sPaging=rxFind(sUrl,"\/((page|show)[0-9]+)(-)?((page|show)?[0-9]+)?/")
If sPaging<>"" Then
sUrl=rxReplace(sUrl,sPaging,"/")
sQuery="@page="&VOV(rxFind(rxFind(sPaging,"page[0-9]+"),rxNumbers),1)&"&"&sQuery
sQuery="@size="&VOV(rxFind(rxFind(sPaging,"show[0-9]+"),rxNumbers),25)&"&"&sQuery
End If
If sUrl="" Then
UpdateCollection j,Request.ServerVariables
DIE "Unable to identify request url"&vbCRLf&vbCRLf&"<br/>"&PrintCol(j)
Else
If IsBack Then
sUrl=rxReplace(sUrl,"\/+","/")
End If
sUrl=rxReplace(sUrl,"/\.(\w{2,4})$",".$1")
End If
End If
If sQuery<>"" Then sQuery=Replace(sQuery&"&","&&","&")
Dim sRef:sRef=REFERRER
If IsFront And(Not IsStaff)Then
If(Not InStr(1,sRef,SERVER_NAME)>0)Then
Dim sFld,rq,rqi
Dim sFlds:sFlds="q,p,ask,searchfor,key,query,search,keyword,keywords,qry,searchitem,kwd,recherche,search_text,search_term,term,terms,qq,qry_str,qu,s,k,t,va"
Dim aFlds:aFlds=Split(sFlds,",")
For Each sFld In aFlds
rqi=rxFind(sRef,"((\?)|(\&))"&sFld&"\=")
If rqi<>"" Then
rq=Mid(sRef,InStr(1,sRef,rqi)+Len(rqi))
If InStr(rq,"&")Then rq=Left(rq,InStr(1,rq,"&")-1)
rq=URLDecode(rq)
Me.Data("refq")=rq:Me.Data("_REFQ")=Me.Data("refq")
Exit For
End If
Next
sRef=rxSelect(sRef,"^[\w\d\.\/\:\-\_]+")
sRef=rxReplace(sRef,"_+W0QQ.*","")
sRef=rxReplace(sRef,"/_+","/")
sRef=rxReplace(sRef,"_+","_")
Dim sESrv
If InStr(1,sRef,"mail")>0 Then
For Each sESrv In Split("mail.yahoo.com,mail.google.com,webmail.aol.com,mail.live.com,webmail.tiscali.co.uk,btinternet.com",",")
sRef=rxReplace(sRef,"http://.*?("&rxEscape(sESrv)&").*","http://$1/")
Next
sRef=rxReplace(sRef,"http://.*?(fsmail)\d+\.("&rxEscape("orange.co.uk")&").*","http://$1.$2/")
End If
Me.Data("ref")=sRef
Me.Data("_REF")=Me.Data("ref")
MySession("ref")=Me.Data("ref")
MySession("refq")=Me.Data("refq")
End If
End If
With Me
.URL=Replace(sUrl,"/.asp",".asp")
.QUERY=sQuery
If Not IsSOAP Then
.FORM=Request.Form
For Each x In Request.Form
If Request.Form(x)="" Then
.Q(x)=""
.Defaults(x)=""
End If
Next
End If
If .FQ("q")<>"" Then
.Q("q0")=.FQ("q")
.Q("q")=Trim(SearchWords(.FQ("q")))
End If
If .FQ("refq")<>"" Then
.Q("refq0")=.FQ("refq")
.Q("refq")=Trim(SearchWords(.FQ("refq")))
End If
Me.SetIDs()
Me.SetRefs()
Me.SetDates()
Me.SetSearchTerm()
End With
End Function
Function SetUrl(ByVal u)
With Me
.URL=u
.SetIDs()
End With
End Function
Function SetDates()
With Me
.AliasD("date:on;start:since,from,after;end:until,before,by")
j("--date-search")=.DateSearch()
j("--name-search")=.NameSearch()
End With
End Function
Function SetRefs()
With Me
Dim utm
For Each utm In Split("source,medium,term,content,campaign",",")
utm="utm_"&p
j(utm)=.FQ(utm)
If IsObject(MySession)Then
If j(utm)<>"" And MySession(utm)="" Then
MySession(utm)=j(utm)
End If
End If
Next
End With
End Function
Function SetSearchTerm()
With Me
If .FQ("q")<>"" Then
DBG "START QUERY > store original q ('"&j("q")&"')"
j("q")=.FQ("q")
j("q-original")=.FQ("q")
Else
DBG "START QUERY > no q"
End If
End With
End Function
Function SetIDs()
With Me
.PATH=Replace(.URL,"/.",".")
.PATHID=rxReplace(.PATH,"^/","")
.PATHID=Left(.PATHID,InStrRev(.PATHID,"/"))
.PATHID=Replace(.PATHID,"/",".")':.PATHID=Replace(.PATHID,".x.",""):.PATHID=Replace(.PATHID,".fwx.back.","")
.PATHID=rxReplace(.PATHID,"^\.|\.$","")
.PAGE=rxSelect(.URL,"[^/]*$")
.PAGEID=Replace(.PAGE,"/",".")':.PAGEID=Replace(.PAGEID,".x.",""):.PAGEID=Replace(.PAGEID,".fwx.back.","")
.PAGEID=rxReplace(.PAGEID,"^\.|\.$","")
If .EXT<>"" Then
.PATHID=rxReplace(.PATHID,"\.("&EXT&")$","")
.PAGEID=rxReplace(.PAGEID,"\.("&EXT&")$","")
Else
.PATHID=rxReplace(.PATHID,"\.(asp|html?)$","")
.PAGEID=rxReplace(.PAGEID,"\.(asp|html?)$","")
End If
.FULLID=Enclose(.PATHID,"",".")&.PAGEID
.FULLID=rxReplace(.FULLID,"^\.|\.$","")
End With
End Function
Public Defaults
Public Function LoadDefaults()
Set Me.Defaults=Dictionary()
If IsPost Then s=Request.Form("z0q")
If s="" Then s=Request.QueryString("z0q")
If s<>"" Then
s=URLDecode(s)
Me.QUERY=s
UpdateCollection Me.Defaults,s
End If
End Function
Public Property Get QueryString()
Dim s,i
s=Request.QueryString
If Left(s,1)="?" Then s=Mid(s,2)
If Left(s,4)="404;" Then
i=InStr(1,s,"?")
If i>0 Then
s=Mid(s,i+1)
Else
s=""
End If
End If
QueryString=s
End Property
Public Property Get String()
String=QString
End Property
Public Property Get QUERY()
QUERY=QString
End Property
Public Function Update(ByVal z)
z=ReAssignQuery(z)
Set QString=New QuickString
Set Me.Data=Dictionary()
Me.QUERY=z
End Function
Public Property Let String(ByVal z)
Me.QUERY=z
End Property
Public Function Add(ByVal z)
Me.QUERY=z
End Function
Public Property Let QUERY(ByVal z)
G_sQueryVars=""
Set QString=New QuickString
Dim i,p,v
For Each y In Split(z,"&")
If y<>"" Then
p=y:v="":i=InStr(1,y,"=")
If i>0 Then
p=Left(y,i-1):v=Mid(y,i+1)
End If
QString.Add p&"="&v&"&"
If InStr(1,p,"%")Or InStr(1,p,"+")Then p=UTF8Decode(p)
If InStr(1,v,"%")Or InStr(1,v,"+")Then v=UTF8Decode(v)
If Me.Data.Exists(p)Then
If Me.Defaults.Exists(p)Then
Me.Defaults.Remove(p)
Me.Data.Remove(p)
Else
Me.Data(p)=Me.Data(p)&", "
End If
End If
Me.Data(p)=Me.Data(p)&v
End If
Next
QString=QString&""
End Property
Public Property Get Q(ByVal fields)
Dim s,o:o=""
If(Not IsSOAP)Then
For Each s In Split(fields,",")
s=TOE(s)
If o="" And s<>"" Then
If Me.Data.Exists(s)Then
o=Me.Data(s)
Else
If Me.Defaults.Exists(s)Then
o=Me.Defaults(s)
End If
End If
End If
Next
End If
Q=o
End Property
Public Property Let Q(ByVal s,ByVal v)
For Each p In Split(s,",")
Me.Data(s)=v
Next
End Property
Public Property Get FORM()
FORM=FString
End Property
Public Property Let FORM(ByVal z)
G_sFormVars=""
z=FString&z
Set Fata=Dictionary()
Set FString=New QuickString
Dim i,p,v
For Each y In Split(z,"&")
If y<>"" Then
p=y:v="y":i=InStr(1,y,"=")
If i>0 Then
p=Left(y,i-1):v=Mid(y,i+1)
End If
FString.Add p&"="&v&"&"
If InStr(1,p,"%")Or InStr(1,p,"+")Then p=UTF8Decode(p)
If InStr(1,v,"%")Or InStr(1,v,"+")Then v=UTF8Decode(v)
If Me.Fata.Exists(p)Then Me.Fata(p)=Me.Fata(p)&", "
Me.Fata(p)=Me.Fata(p)&v
End If
Next
FString=FString&""
End Property
Public Property Get F(ByVal fields)
Dim s,o:o=""
If(Not IsSOAP)Then
For Each s In Split(fields,",")
s=TOE(s)
If o="" And s<>"" Then
If Me.Fata.Exists(s)Then
o=Me.Fata(s)
End If
End If
Next
End If
F=o
End Property
Public Property Let F(ByVal s,ByVal v)
For Each p In Split(s,",")
Me.Fata(s)=v
Next
End Property
Public Property Get FQ(ByVal fields)
Dim s,o:o=""
For Each s In Split(fields,",")
If o="" Then
s=TOE(s)
If s<>"" Then
If(o="")And Me.Fata.Exists(s)Then o=Me.Fata(s)
If(o="")And Me.Data.Exists(s)Then o=Me.Data(s)
If(o="")And Me.Defaults.Exists(s)Then o=Me.Defaults(s)
End If
End If
Next
FQ=o
End Property
Public Property Get QF(ByVal fields)
Dim s,o:o=""
For Each s In Split(fields,",")
If o="" Then
s=TOE(s)
If s<>"" Then
If(o="")And Me.Data.Exists(s)Then o=Me.Data(s)
If(o="")And Me.Fata.Exists(s)Then o=Me.Fata(s)
If(o="")And Me.Defaults.Exists(s)Then o=Me.Defaults(s)
End If
End If
Next
QF=o
End Property
Public Sub Alias(ByVal s)
Dim a,v,b
With Me
For Each a In Split(s,";")
a=Split(a,":")
v=.FQ(a(0)&","&a(1))
If v<>"" Then
.Q(a(0))=v
End If
Next
End With
End Sub
Public Sub AliasD(ByVal s)
Dim a,v,b,f,inQ,inF
With Me
For Each a In Split(s,";")
inQ=False
inF=False
a=Split(a,":")
v=.DData(a(0)&","&a(1))
If v<>"" Then
For Each f In Split(a(0)&","&a(1),",")
If .Data.Exists(f)Then
If .Data(f)<>"" Then
inQ=True
End If
.Data.Remove(f)
End If
If .Fata.Exists(f)Then
If .Fata(f)<>"" Then
inF=True
End If
.Fata.Remove(f)
End If
Next
If inQ Then
.Data(a(0))=v
Else
If .Data.Exists(a(0))Then .Data.Remove(a(0))
End If
If inF Then
.Fata(a(0))=v
Else
If .Fata.Exists(a(0))Then .Fata.Remove(a(0))
End If
End If
Next
End With
End Sub
Function ParsedData(ByVal f)
ParsedData=ParseData(f)
End Function
Function ParseData(ByVal f)
ParseData=Me.FQ(f)
End Function
Function TData(ByVal fs)
Dim f,o:o=""
For Each f In Split(fs,",")
o=Me.ParsedData(f)
If o<>"" Then
TData=o
Exit Function
End If
Next
TData=""
End Function
Function DData(ByVal fs)
Dim f,d,o:o=""
For Each f In Split(fs,",")
If d="" Then
o=Me.ParsedData(f)
If o<>"" Then
If IsDate(o)Then
d=o
End If
If rxTest(o,"^[\-0-9]+$")Then
If Len(o)>10 Then
d=DateAdd("s",VOZ(o),"1 Jan 1970 00:00")
ElseIf Abs(VOZ(o))<=365 Then
d=DateAdd("d",VOZ(o),Date())
End If
End If
DBG "QUERY > parse date input '"&f&"'=DateFromString('"&o&"')"
d=NewLocalDate(o)
If d<>"" Then
DBG "QUERY > parse date input '"&f&"'=DateFromString('"&o&"') = '"&d&"'"
Else
DBG "QUERY > parse date input '"&f&"'=DateFromString('"&o&"') FAILED!"
o=""
End If
End If
End If
Next
If d<>"" Then
d=FmtDate(d,"%Y-%M-%D")
End If
DData=d
End Function
Function NData(ByVal fs)
Dim f,o:o=0
For Each f In Split(fs,",")
o=VOZ(Me.ParsedData(f))
If o<>0 Then
NData=o
Exit Function
End If
Next
NData=0
End Function
Function BData(ByVal fs)
BData=IsTrue(Me.TData(fs))
End Function
Function YN(ByVal fs)
YN=CBool(Me.BData(fs))
End Function
Function Has(ByVal fs)
Has=CBool(Me.TData(fs)<>"")
End Function
Function DateSQL():DateSQL=TOT(Enclose(Me.DateSearch(),"(",")"),"1=1"):End Function
Function NotDateSQL():NotDateSQL=TOT(Enclose(Me.DateSearch(),"(NOT ",")"),"1=1"):End Function
Function DateSQLFor(ByVal t):DateSQLFor=rxReplace(Me.DateSQL(),"\b(t)\.",t&"."):End Function
Function NotDateSQLFor(ByVal t):NotDateSQLFor=rxReplace(Me.NotDateSQL(),"\b(t)\.",t&"."):End Function
Function DateSearch()
If Me.Data("date-search")="" Then
If Me.FQ("date,start,end")<>"" Then
Dim f,o
Dim timezone,dd,ds,de
timezone=MyTimezone
Set f=New fwxSearch
f.Reset
f.DefaultConfig="m-a"
f.AddFilter fwxDateSearch,"m-a"
o=f.SQL
o=Enclose(o,"(",")")
Set f=Nothing
Me.Data("date-search")=o
End If
If Me.FQ("date")<>"" Then
j("date-search-label")=j("date-search-label")&" for "&FmtDate(Me.FQ("date"),"A b dO Y")
ElseIf Me.FQ("start,end")<>"" Then
j("date-search-label")=j("date-search-label")&Enclose(FmtDate(Me.FQ("start"),"A b dO Y")," from ","")
j("date-search-label")=j("date-search-label")&Enclose(FmtDate(Me.FQ("end"),"A b dO Y")," until ","")
End If
j("date-search-or-ever")=TOT(Me.Data("date-search"),"99=99/*any date*/")
j("date-search-or-today")=TOT(Me.Data("date-search"),"t.date=getdate()")
End If
DateSearch=Me.Data("date-search")
End Function
Function StringSearch()
StringSearch=Me.NameSearch()
End Function
Function NameSearch()
If Me.Data("name-search")="" Then
If Me.FQ("q")<>"" Then
Dim f,o
Set f=New fwxSearch
f.Reset
f.DefaultConfig="m-a"
f.AddFilter "t.name:text:s-q","m-a"
o=f.SQL
o=Enclose(o,"(",")")
Set f=Nothing
Me.Data("name-search")=o
End If
End If
NameSearch=Me.Data("name-search")
End Function
End Class
Function H404__XSL():H404__XSL=H404__XML:End Function
Function H404__XML()
If IsTest Then
PrivateCacheResponse 0
Else
PublicCacheResponse 525600
End If
Dim url:url=Query.RAWURL
Dim folder,temp_file,last_modified
Dim file,file2,found
Dim gzip,pack,code_static,code_packed
Dim domain_folders,system_folders
domain_folders="/U/ "&Replace(Replace(FE__TMPL_Folders,"/fe",""),";"," ")&" "&TMPL_FOLDER&" /"
system_folders="/fwx/ /fwx/@CODE/"' /fwx/inc/"
If False Then
ElseIf rxTest(url,"^domain/")Then
url=rxReplace(url,"^domain/","")
folders=domain_folders
ElseIf rxTest(url,"^system/")Then
url=rxReplace(url,"^system/","")
folders=system_folders
Else
folders=domain_folders&" "&system_folders
End If
folders=Trim(folders)
DBG "H404__CSS:"&url
If rxTest(url,"^/(((x|U|@|fwx)/)|(@\.))")Then
url=rxReplace(url,"^/((x|U|@|fwx)/)?","")
file=url
gzip=rxTest(file,"^(compress|g?zipp?)(ed)?/")
file=rxReplace(file,"^(compress|g?zipp?)(ed)?/","")
pack=rxTest(file,"^(pack)(er|ed)?/")
file=rxReplace(file,"^(pack)(er|ed)?/","")
found=FWX.Setting("RESOURCE:"&file)
If found="" Then
For Each folder In Split(folders," ")
If found="" Then
DBG "H404__XML: "&folder&file&vbCRLf
If FileExists(folder&file)Then
found=folder&file
Else
file2=rxReplace(file,"^(x|U|@|fwx)/","")
If file2<>file Then
DBG "H404__XML: "&folder&file2&vbCRLf
If FileExists(folder&file2)Then found=folder&file2
End If
End If
End If
Next
FWX.Setting("RESOURCE:"&file)=found
End If
found=TranslatedPath(found)
DBG "H404__XML: found translated= '"&found&"'"
If found<>"" Then
CacheControl LastModified(found)
If pack Then
Server.ScriptTimeout=60
code_static=ReadFile(found)
code_packed=CrunchXML(code_static)
If code_packed<>"" Then
If gzip Or USE_GZIP Or IsTrue(FWX.Setting("gzip"))Then
GZipContent code_packed
End If
DIE code_packed
End If
End If
If gzip Or USE_GZIP Or IsTrue(FWX.Setting("gzip"))Then
GZipUrl found
End If
Transfer found
End If
If IsFront Then
End If
End If
End Function
Function H404_bo_js()
PermanentRedirect "/fwx/fwx.bo.js"
End Function
Function H404_fe_js()
PermanentRedirect "/fwx/fwx.fe.js"
End Function
Function H404__JS():H404__JS=H404__CSS:End Function
Function H404__HTC():H404__HTC=H404__CSS:End Function
Function H404__CSS()
PublicCacheResponse 525600
Dim url:url=Query.RAWURL
Dim folder,temp_file,last_modified
Dim file,file2,found
Dim gzip,pack,code_static,code_packed
If url="/ga.js" Then
Call GoogleAnalyticsCode()
RE
End If
If rxTest(url,"^/(static|source)/(fe|bo)(\-.+)?\.(js|css)$")Then
url=rxReplace(url,"^/(static|source)/(fe|bo)(\-.+)?\.(js|css)","/static/$2.$4")
End If
If rxTest(url,"/(fwx|@)/(fe|bo)\.(js|css)$")Then
Transfer rxReplace(url,"/(fwx|@)/(fe|bo)\.(js|css)$","/fwx/@@SRC/@COMPILER/ASSEMBLY/@CODE/$2/$2.$3.asp")
End If
found=FWX.Setting("RESOURCE:"&url)
If found="" Then
If rxTest(url,"^/(((x|U|static|@|fwx)/)|(@\.))")Then
file=url
file=rxReplace(file,"^/((x|U|static|@|fwx)/)?","")
gzip=rxTest(file,"^(compress|g?zipp?)(ed)?/")
file=rxReplace(file,"^(compress|g?zipp?)(ed)?/","")
pack=rxTest(file,"^(pack)(er|ed)?/")
file=rxReplace(file,"^(pack)(er|ed)?/","")
Dim domain_folders,system_folders
domain_folders="/U/ "&Replace(Replace(FE__TMPL_Folders,"/fe",""),";"," ")&" "&TMPL_FOLDER&" /"
system_folders="/fwx/ /fwx/@CODE/ /fwx/inc/ /fwx/js/"&IIf(rxTest(file,"jquery")," /fwx/js/jquery/","")
If False Then
ElseIf rxTest(file,"^domain/")Then
file=rxReplace(file,"^domain/","")
folders=domain_folders
ElseIf rxTest(file,"^system/")Then
file=rxReplace(file,"^system/","")
folders=system_folders
Else
folders=domain_folders&" "&system_folders
End If
folders=Trim(folders)
DBG "H404__CSS:file="&file
DBG "H404__CSS:folders="&folders
For Each folder In Split(folders," ")
If found="" Then
DBG "H404__CSS/H404__JS: "&folder&file&vbCRLf
If FileExists(folder&file)Then
found=folder&file
Else
file2=rxReplace(file,"^(x|U|@|fwx)/","")
If file2<>file Then
DBG "H404__CSS/H404__JS: "&folder&file2&vbCRLf
If FileExists(folder&file2)Then found=folder&file2
End If
End If
End If
Next
FWX.Setting("RESOURCE:"&file)=found
End If
DBG "H404__CSS/H404__JS: FOUND? "&IIf(found<>"",found,"NOT FOUND")&" (pack="&pack&")"
If found<>"" Then
found=TranslatedPath(found)
DBG "H404__CSS/H404__JS: found translated= "&found&""
End If
If found<>"" Then
DBG "H404__CSS/H404__JS: FOUND FILE "&found&""
CacheControl LastModified(found)
If pack Then
Server.ScriptTimeout=5
DBG "H404__CSS/H404__JS: PACK FILE "&found&""
code_static=ReadFile(found)
If Query.EXT="css" Then
code_packed=ObfuscateCSS(code_static)
ElseIf Query.EXT="js" Then
code_packed=JSObfuscator.Pack(found)
End If
If code_packed<>"" Then
If gzip Or USE_GZIP Or IsTrue(FWX.Setting("gzip"))Then
GZipContent code_packed
End If
DIE code_packed
End If
End If
If gzip Or USE_GZIP Or IsTrue(FWX.Setting("gzip"))Then
DBG "H404__CSS/H404__JS: TRY GZIP FILE "&found&""
GZipUrl found
End If
DBG "H404__CSS/H404__JS: TRANSFER TO "&found&""
Transfer found
End If
If rxTest(file,"(bo|fe)\.(js|css)$")Then
Dim compiled_file
compiled_file="/fwx/fwx."&rxSelect(file,"(bo|fe)\.(js|css)$")
DBG "H404__CSS/H404__JS: TRY TRANSFER TO "&compiled_file&""
If FileExists(compiled)Then
Transfer compiled_file
DIE ""
End If
End If
If Not InStr(1,url,"@@SRC")>0 Then
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source)"
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) ::A:: "&file
Dim base:base="/fwx/@@SRC/@COMPILER/ASSEMBLY/"
file=rxReplace(TOT(file,url),"^/?(fwx/)?(@CODE/)?((fe|bo)/)?","")
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) ::B:: "&file
Dim source_url:source_url=""
Select Case LCase(rxReplace(file,"\.(js|css)$",""))
Case "@","fwx":
source_url=base&"fwx."&Query.EXT&".asp"
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > core = source_url= "&source_url
Transfer source_url
Case "fe":
source_url=base&"fwx.fe."&Query.EXT&".asp"
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > fe = source_url= "&source_url
Transfer source_url
Case "bo":
source_url=base&"fwx.bo."&Query.EXT&".asp"
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > bo = source_url= "&source_url
Transfer source_url
Case Else:
file=file&IIf(rxTest(file,"^(css|js)"),"",".asp")
If InStr(1,file,"@CODE")>0 Then
source_url=base&rxReplace(file,"^@CODE/(fe|bo)/","fwx.")
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > other X = source_url= "&source_url
Transfer source_url
End If
source_url=base&Query.EXT&"/"&file
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > other A = source_url= "&source_url
Transfer source_url
source_url=base&file
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > other B = source_url= "&source_url
Transfer source_url
source_url=base&"@CODE/"&Query.EXT&"/"&file
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > other C = source_url= "&source_url
Transfer source_url
source_url=base&"@CODE/"&file
DBG "H404__CSS/H404__JS: TRY TRANSFER TO COMPILER (direct from source) > other D = source_url= "&source_url
Transfer source_url
End Select
End If
If IsFront Then
End If
End If
End Function
Function H404__JPEG():H404__JPEG=H404__IMG():End Function
Function H404__PNG():H404__PNG=H404__IMG():End Function
Function H404__BMP():H404__BMP=H404__IMG():End Function
Function H404__GIF():H404__GIF=H404__IMG():End Function
Function H404__JPE():H404__JPE=H404__IMG():End Function
Function H404__JPG():H404__JPG=H404__IMG():End Function
Function H404__IMG()
PublicCacheResponse 525600
Dim url
url=Query.RAWURL
DBG "H404__IMG: url="&url&""
found=ENV_Resource_Image(url)
DBG "H404__IMG: found="&found&""
If found="" Then
If IsTest Then
If DBG.Show Then
DIE "would have redirected to www - http://"&MyWWWDomain&url
Else
Redirect "http://"&MyWWWDomain&url 'DBG.HTTP404 "Image '"&url&"' not found! ("&TMPL_FOLDER&")"
End If
Else
DBG.HTTP404 "Image '"&url&"' not found! ("&TMPL_FOLDER&")"
End If
End If
DBG "H404__IMG: cache control for "&found&""
CacheControl LastModified(found)
If Query.RAWURL<>found Then
DBG "H404__IMG: virtual folder resolved, image found in: "&found
DBG "H404__IMG: Pretend the request came from "&found&""
Query.RAWURL=found
End If
DBG "H404__IMG: H404__IMG__Virtualize("&found&")"
Call H404__IMG__Virtualize()
End Function
Function H404__IMG__Virtualize()
Dim sUrl:sUrl=Query.RAWURL
Dim rxResize:rxResize="/([0-9]+)x([0-9]+)"
Dim rxCmsImg:rxCmsImg="/(thumb|image|main|icon|i1|i2|i3)\."
Dim rxSeoTxt:rxSeoTxt="\(.+\-\)|\(\-.+\)|/fwx-[^\.]+|/--[^\.]+"
Dim rsShrtCt:rsShrtCt="^/(x|i|U|@)/"
Dim sImgFile,sTrgFile,bData
Dim oImager,iWidth,iHeight,bDistort
Dim sPageUrl,sImgType
Dim rImg,sqlImg
Dim aMatches
iWidth=0:iHeight=0:bDistort=False
DBG "H404__IMG__Virtualize: Prepare - sUrl = '"&sUrl&"'"&vbCRLf
DBG "H404__IMG__Virtualize: Prepare - SEO content = "&rxSelect(sUrl,rxSeoTxt)&vbCRLf
DBG "H404__IMG__Virtualize: Prepare - Shortcut prefix = "&rxSelect(sUrl,rsShrtCt)&vbCRLf
DBG "H404__IMG__Virtualize: Prepare - Resize parameters = "&rxSelect(sUrl,rxResize)&vbCRLf
If 1=0 Then
ElseIf rxTest(sUrl,rxResize)Then
DBG "H404__IMG__Virtualize: *** RESIZE HAS BEEN REQUESTED ***"&vbCRLf
DBG "H404__IMG__Virtualize: *** shorcuts cannot be used when resizing is requested"&vbCRLf
Else
DBG "H404__IMG__Virtualize: *** A FIXED IMAGE IN FSO HAS BEEN REQUESTED ***"&vbCRLf
End If
sImgFile=rxReplace(sUrl,rxSeoTxt,"")
DBG "H404__IMG__Virtualize: START! sImgFile = '"&sImgFile&"'"&vbCRLf
If rxTest(sImgFile,rxResize)Then
DBG "H404__IMG__Virtualize: RESIZE"&vbCRLf
Set aMatches=rxMatches(sImgFile,rxResize)
If aMatches.Count>1 Then
DBG.FatalError "ERROR: Too many resizing matches"
End If
iWidth=VOZ(Mid(aMatches(0),2,InStr(1,aMatches(0),"x")-2))
iHeight=VOZ(Mid(aMatches(0),InStr(1,aMatches(0),"x")+1))
Set aMatches=Nothing
sImgFile=Replace(r.Replace(sImgFile,"/"),"//","/")
DBG "H404__IMG__Virtualize: RESIZE: sImgFile = '"&sImgFile&"'"&vbCRLf
r.Pattern="\/(distort|distorted|strech|streched|stretch|stretched)\/"
If r.Test(sImgFile)Then
sImgFile=r.Replace(sImgFile,"/")
bDistort=True
End If
DBG "H404__IMG__Virtualize: RESIZE: bDistort = '"&bDistort&"'"&vbCRLf
sImgFile=ENV_Resource_Image(sImgFile)
DBG "H404__IMG__Virtualize: RESOLVED: sImgFile = '"&sImgFile&"'"&vbCRLf
End If
If sImgFile="" Then
DBG.HTTP404 "Image not found"
End If
If iWidth>0 Then
DBG "H404__IMG__Virtualize: Resizing required"&vbCRLf
iWidth=VOZ(iWidth):iHeight=VOZ(iHeight)
DBG "H404__IMG__Virtualize: Resizing target: "&iWidth&"w x "&iHeight&"h"&vbCRLf
If Len(bData)>0 Or sTrgFile<>"" Then
Else
DBG "H404__IMG__Virtualize: PHP: Try..."
sUrlPath=MyProtocol&"://"&SERVER_NAME&sImgFile
BuildPath "/$$TEMP/$IMAGES/phpThumb/"
Dim phpThumb:phpThumb="/fwx/php/images/phpThumb.php?w="&iWidth&"&h="&iHeight&"&src="&URLEncode(sUrlPath)&"&"
DBG "H404__IMG__Virtualize: PHP: Calling script: <a href='"&phpThumb&"'>URL</a>"
bData=GetResponseBody(phpThumb)
DBG "H404__IMG__Virtualize: PHP: Script response: "&Len(bData)&" bytes"
If Left(BinaryToText(bData),3)="PHP" Then
DBG "H404__IMG__Virtualize: PHP: ERROR: "&bData&""
Else
End If
End If
If Len(bData)>0 Or sTrgFile<>"" Then
Else
Dim sResizeScript:sResizeScript="/fwx/vb/Images/Resizer.asp?w="&iWidth&"&h="&iHeight&"&d="&IIf(bDistort,"y","")&"&i="&sImgFile&"&p=y&stamp="&RemoveSymbols(Now())&"&"
If Query.FQ("t")<>"" Then sResizeScript=sResizeScript&"t=y&"
DBG "H404__IMG__Virtualize: ASP: Calling script: <a href='"&sResizeScript&"'>URL</a>"
sTrgFile=GetResponse(sResizeScript)
DBG "H404__IMG__Virtualize: ASP: Script response: "&sTrgFile&""
If Query.FQ("t")<>"" And InStr(1,sTrgFile,"DONE!")>0 Then
sTrgFile=rxReplace(rxSelect(sTrgFile,"[\r\n].*$"),"[\r\n]+","")
Else
If InStr(1,sTrgFile,"<")>0 Then
sTrgFile=""
End If
End If
If sTrgFile<>"" Then
DBG "H404__IMG__Virtualize: ASP: Resizing worked - <a href='"&sTrgFile&"'>URL</a>"
Else
DBG "H404__IMG__Virtualize: ASP: Resizing failed - <a href='"&sResizeScript&"'>URL</a>"
End If
End If
End If
If(DBG.Show Or Query.FQ("t")<>"")Then
DBG "H404__IMG__Virtualize: CONTENT TYPE > image/"&rxReplace(rxSelect(sImgFile,"\w{3,4}$"),"jp(g|ge|eg)","jpeg")
DBG "H404__IMG__Virtualize: END > "&sImgFile
RHTM
Else
Response.Clear()
Response.ContentType="image/"&rxReplace(rxSelect(sImgFile,"\w{3,4}$"),"jpg?e?","jpeg")
If IsTest Then
If Len(bData)>0 Then
Else
sTrgFile=TOT(sTrgFile,sImgFile)
If sTrgFile<>"" Then
If(USE_GZIP Or IsTrue(FWX.Setting("gzip")))Then GZipUrl sTrgFile
bData=ReadBinary(sTrgFile)
End If
End If
End If
If Len(bData)>0 Then
FWX_ResponseStatus 200
Response.BinaryWrite bData
Else
sTrgFile=TOT(sTrgFile,sImgFile)
If sTrgFile<>"" And sTrgFile<>Query.URL Then
Transfer sTrgFile
End If
FWX_ResponseStatus "404 Not Found"
End If
End If
Response.End()
End Function
Function H404__Commands()
Dim sUrl:sUrl=Query.RAWURL
If Not DBG.Show Then Response.ExpiresAbsolute=Date()-1
sUrl=rxSelect(sUrl,"\;(.+)$")
sUrl=rxSelect(sUrl,"\w+")
If EXE("H404__CMD_"&sUrl&"")Then RE
Dim ACTION_URL:ACTION_URL="/fwx/@.asp?|do="&sUrl
Response.Write "<a href='"&ACTION_URL&"'><strong>H404__Commands - FWX Action &raquo;</strong></a>"
If IsLive Then
Response.Write "<scr"&"ipt>window.location='"&ACTION_URL&"'</scr"&"ipt>"
End If
Response.End()
End Function
Function H404__CMD__END():H404__CMD__END=True
Dim xDomain_gave,xDomain_want
If xDomain_gave="" Then xDomain_gave=Query.FQ("Access-Control-Allow-Origin")
If xDomain_gave="" Then xDomain_gave=Request.ServerVariables("Access-Control")
If xDomain_gave<>"" Then 'Query.FQ("Access-Control-Allow-Origin")<>"" Then
xDomain_want=MD5(FmtDate(Date(),"%y%M%D"))
If xDomain_want=xDomain_gave Then
FWX_AddHeader "Access-Control-Allow-Origin","*" '"."&MyDomain
Else
RC
DIE "xDomain_want="&xDomain_want&". xDomain_gave="&xDomain_gave&". "&Request.ServerVariables
End If
End If
If IsAjax Then
ElseIf Query.FQ("stay")<>"" Then
Else
Redirect j("REDIRECT")
j("REDIRECT")=Query.RAWURL
j("REDIRECT")=rxReplace(j("REDIRECT"),"/;\w+/?$","")
j("REDIRECT")=rxReplace(j("REDIRECT"),"/+$","")
j("REDIRECT")=TOT(j("REDIRECT"),"/")
If IsHuman Then
RW "<br/>"
RW "<a href='"&j("REDIRECT")&"'>CONTINUE &raquo;</a>"
RW "<script>if(true) top.location='"&j("REDIRECT")&"';</script>"
Else
RC
PermanentRedirect j("REDIRECT")
End If
End If
RE
End Function
Function H404__CMD_comp():H404__CMD_comp=H404__CMD_compiler:End Function
Function H404__CMD_compile():H404__CMD_compile=H404__CMD_compiler:End Function
Function H404__CMD_compiler():H404__CMD_compiler=True
PermanentRedirect "/fwx/@@SRC/@COMPILER/"
End Function
Function H404__CMD_fonts():H404__CMD_fonts=H404__CMD_font:End Function
Function H404__CMD_font():H404__CMD_font=True
PermanentRedirect "/fwx/inc/fontello/demo.html"
End Function
Function H404__CMD_cy():H404__CMD_cy=H404__CMD_con:End Function
Function H404__CMD_c1():H404__CMD_c1=H404__CMD_con:End Function
Function H404__CMD_con():H404__CMD_con=True
MyCookies.Global("NoCache.Flag")=""
RW "<b class=Yes>ON</b>"
CALL H404__CMD__END()
End Function
Function H404__CMD_cn():H404__CMD_cn=H404__CMD_coff:End Function
Function H404__CMD_c0():H404__CMD_c0=H404__CMD_coff:End Function
Function H404__CMD_coff():H404__CMD_coff=True
MyCookies.Global("NoCache.Flag")="y"
RW "<b class=No>OFF</b>"
CALL H404__CMD__END()
End Function
Function H404__CMD_ct():H404__CMD_ct=True
If MyCookies.Global("NoCache.Flag")="y" Then
CALL H404__CMD_con()
Else
CALL H404__CMD_coff()
End If
End Function
Function H404__CMD_cs()
Query.Fata("stay")="y"
H404__CMD_cs=H404__CMD_c()
End Function
Function H404__CMD_Ly():H404__CMD_Ly=H404__CMD_Lon:End Function
Function H404__CMD_L1():H404__CMD_L1=H404__CMD_Lon:End Function
Function H404__CMD_Lon():H404__CMD_Lon=True
MyCookies.Global("LiveOnLocal")="y"
RW "LiveOnLocal=<b class=Yes>ON</b>"
CALL H404__CMD__END()
End Function
Function H404__CMD_Ln():H404__CMD_Ln=H404__CMD_Loff:End Function
Function H404__CMD_L0():H404__CMD_L0=H404__CMD_Loff:End Function
Function H404__CMD_Loff():H404__CMD_Loff=True
MyCookies.Global("LiveOnLocal")=""
RW "LiveOnLocal=<b class=Yes>OFF</b>"
CALL H404__CMD__END()
End Function
Function H404__CMD_Lt():H404__CMD_Lt=True
If MyCookies.Global("LiveOnLocal")="y" Then
CALL H404__CMD_Loff()
Else
CALL H404__CMD_Lon()
End If
End Function
Function H404__CMD_s():H404__CMD_s=True
If IsObject(Session)Then
Session.Contents.RemoveAll():Session.Abandon()
End If
For Each c In Request.Cookies
Response.Cookies(c)=""
Response.Cookies(c).Expires=Now()-1
Next
RW "Session ended"
CALL H404__CMD__END()
End Function
Function H404__CMD_ss()
Query.Fata("stay")="y"
H404__CMD_ss=H404__CMD_s()
End Function
Function H404__CMD__EDITION():H404__CMD__EDITION=True
RW "FWX.EDITION="&FWX.Setting("EDITION")&"<br/>"&vbCRLf
RW "FWX.REVISION="&FWX.Setting("REVISION")&"<br/>"&vbCRLf
End Function
Function H404__CMD_edition():H404__CMD_edition=True
CALL H404__CMD__EDITION()
RE
End Function
Function H404__CMD_cf():H404__CMD_cf=True
If Not IsStaffPC Then Exit Function
Query.Fata("stay")="y"
RW "Cloudflare purge: "&CloudFlarePurge()
DIE ""
End Function
Function H404__CMD_cfon():H404__CMD_cfon=H404__CMD_cf1:End Function
Function H404__CMD_cf1():H404__CMD_cf1=True
If Not IsStaffPC Then Exit Function
Query.Fata("stay")="y"
RW "Cloudflare ON: "&CloudFlare_DevMode_OFF()
DIE ""
End Function
Function H404__CMD_cfoff():H404__CMD_cfoff=H404__CMD_cf0:End Function
Function H404__CMD_cf0():H404__CMD_cf0=True
If Not IsStaffPC Then Exit Function
Query.Fata("stay")="y"
RW "Cloudflare OFF: "&CloudFlare_DevMode_ON()
DIE ""
End Function
Function H404__CMD_c():H404__CMD_c=True
Dim loc,ftr:ftr=""
For Each loc In Split(TOT(Query.FQ("location,map"),"0"),",")
loc=VOZ(loc)
If loc>0 Then
ftr=ftr&"(name LIKE '%{"&loc&"}%') OR"
End If
Next:ftr=TOT(ftr,"1=1 OR")&" 1=0"
Call ExeSQL("TRUNCATE TABLE tblAPP__CACHE")'--DELETE FROM tblAPP__CACHE WHERE site="& MySite &" AND(("& ftr &")OR(name NOT LIKE 'I%'))")
Application.Contents.RemoveAll()
If IsAdmin Then FWX.KillEdition()
FWX.Reboot()
RW "Cache cleared"&vbCRLf
CALL H404__CMD__END()
End Function
Function H404__CMD_p():H404__CMD_p=H404__CMD_reboot:End Function
Function H404__CMD_init():H404__CMD_init=H404__CMD_reboot:End Function
Function H404__CMD_purge():H404__CMD_purge=H404__CMD_reboot:End Function
Function H404__CMD_reset():H404__CMD_reset=H404__CMD_reboot:End Function
Function H404__CMD_golive():H404__CMD_golive=H404__CMD_reboot:End Function
Function H404__CMD_reboot():H404__CMD_reboot=True
DeleteFile "/$DOMAINCACHE/TMPL/RESOLVED-FE.txt"
DeleteFile "/$DOMAINCACHE/TMPL/RESOLVED-BO.txt"
DeleteFile "/$DOMAINCACHE/TMPL/RESOLVED-BO-FORM.txt"
If IsObject(MyCookies)Then MyCookies.Global("NoCache.Flag")="y"
Dim sd,si,sr
sd=Query.FQ("subdomain")
If sd<>"" Then
For Each si In Split(sd,",")
sr=GetResponse("http://"&si&"."&MyDomain&"/;p")
Next
End If
If IsLive Then RW "Cloudflare purge: "&CloudFlarePurge()&"<br/>"&vbCRLf
If IsLive Then RW "Cloudflare DEV mode ON: "&CloudFlare_DevMode_ON()&"<br/>"&vbCRLf
Application.Contents.RemoveAll()
FWX.KillEdition()
FWX.Reboot()
RW "Cache cleared and application restarted"&vbCRLf
CALL H404__CMD__END()
End Function
Function H404__CMD_ps()
Query.Fata("stay")="y"
H404__CMD__PS=H404__CMD_reboot()
End Function
Function H404__CMD_k():H404__CMD_k=True
Redirect rxReplace(Query.RAWURL,"/;k/?$","/;s/;r/;c/;purge")
End Function
Function H404__CMD_r():H404__CMD_r=True
Application.Contents.RemoveAll()
If IsAdmin Then FWX.KillEdition()
FWX.Reboot()
RW "Reset"&vbCRLf
CALL H404__CMD__END()
End Function
Function H404__CMD_staff():H404__CMD_staff=True
RW "Staff flag ON"&vbCRLf
MyCookies.Global("IsStaff.Flag")="y"
MyCookies.Global("NoCache.Flag")="y"
Call H404__CMD__END()
End Function
Function H404__CMD_session():H404__CMD_session=True
rw "### IsLive="&IsLive&vbCRLf
rw "### IsStaff="&IsStaff&vbCRLf
rw "### IsAdmin="&IsAdmin&vbCRLf
rw "### MyCredentials="&MyCredentials&vbCRLf
If IsLive And Not(IsTest Or IsAdmin Or IsStaff)Then RE
rw vbCRLf&"### Session.Contents ###"&vbCRLf
Call H404__CMD__debug(Session.Contents)
End Function
Function H404__CMD_server():H404__CMD_server=True
rw vbCRLf&"### Request.ServerVariables ###"&vbCRLf
Call H404__CMD__debug(Request.ServerVariables)
End Function
Function H404__CMD_app():H404__CMD_app=True
rw vbCRLf&"### Application.Contents ###"&vbCRLf
If IsTest Then Call H404__CMD__debug(Application.Contents)
End Function
Function H404__CMD_cookies():H404__CMD_cookies=True
rw vbCRLf&"### Request.Cookies ###"&vbCRLf
Call H404__CMD__debug(Request.Cookies)
End Function
Function H404__CMD_debug():H404__CMD_debug=True
If IsLive And Not(IsAdmin Or IsStaff)Then RE
Call H404__CMD_session()
Call H404__CMD_app()
Call H404__CMD_cookies()
End Function
Function H404__CMD__debug(ByRef c)
rTXT
Dim s,n
For Each s In c
n=n+Len(c(s))
If IsLocal Then
rw s&"("&Len(c(s))&"): "&(c(s))&vbCRLf
Else
rw s&"("&Len(c(s))&"): "&(Left(c(s),30))&vbCRLf
End If
Next
rw s&"TOTAL DATA: "&n&" bytes"&vbCRLf
End Function
Function H404__CMD_wait():H404__CMD_wait=True
Response.ExpiresAbsolute=Date()+365
Response.CacheControl="public"
Response.Buffer=False
RC
RW "<html><head><title>Please wait...</title><link rel='stylesheet' href='/@/@.css' type='text/css'><script type='text/javascript' src='/x/js/base.js' language='Javascript'></script></head><body>"
RW "<div class='P3 Yes'><strong>Please wait!</strong></div>"
RW "<div class='P3 Disabled'>Loading...</div>"
RW "</body></html>"
RE
End Function
Function H404__CMD_fxl():H404__CMD_fxl=True
Call ExeSQL("EXEC dbo.[pcdSITE__Link_Fix] %SITE")
RW "Links in site "&MySite&" rebuilt on "&Now()&""
CALL H404__CMD__END()
End Function
Function H404__CMD_v():H404__CMD_v=True
Dim tRS:Set tRS=Recordset()
Call OpenRS(tRS,"EXEC dbo.[pcdLIVE_Visitor_Count] %SITE")
Dim o:o=tRS("active")
If IsEmpty(o)Then o="!ERROR!"
RW o
Set tRS=Nothing
RE
End Function
Function H404__CMD_():H404__CMD_=True
Query.URL=rxReplace(Query.URL,"\/\;[^/]*\/","/")
RW "<strong>Select a special command from the list below:</strong>"
RW "<br/>"
RW "<br/>"
RW "<a href="""&Query.URL&";c/"">Clear cache</a>"
RW " - "
RW "<a href="""&Query.URL&";ct/"">Toggle cache</a>"
RW " - "
RW "<a href="""&Query.URL&";fxl/"">Fix links</a>"
RW "<br/>"
RW "<br/>"
RW "<a href="""&Query.URL&";s/"">End session</a>"
RW " - "
RW "<a href="""&Query.URL&";r/"">Reset application</a>"
RW "<br/>"
RW "<br/>"
RW "<a href="""&Query.URL&";config/"">Configuration Utility</a>"
RW "<br/>"
RW "<br/>"
RW "<a href="""&Query.URL&""">&laquo; GO BACK</a>"
RE
End Function
Function H404_cookies()
RTXT
Dim k,v,n,e
e="!" ' label for "empty" value
n=Query.FQ("redirect,next")
k=TOT(Query.FQ("key,name"),rxReplace(Query.URL,"^/cookies[^/]*/?|/$",""))
v=Query.FQ("update,value,val,set,u,"&k)
If UBound(Split(k,"="))=1 Then
v=URLDecode(TOT(rxSelect(k,"[^\=]*$"),e))
k=rxSelect(k,"^[^\=]*")
End If
If k<>"" Then
If v<>"" Then
If v=e Then v=""
MyCookies.Global(k)=v
Redirect n
RW "{'"&k&"':'"&EncodeJS(v)&"'}"
Else
v=MyCookies.Global(k)
RW v&""
End If
Else
Select Case EXT
Case "js"
RW "data={"&vbCRLf&DIC2JS(Request.Cookies)&vbCRLf&"}"
Case "xml"
RW "<cookies>"&vbCRLf&DIC2XML(Request.Cookies)&vbCRLf&"</cookies>"
Case Else
RW "// LIST of cookies in use by '"&SERVER_NAME&"'"&vbCRLf
RW vbCRLf
For Each c In Request.Cookies
RW c&"="&MyCookies(c)&vbCRLf
Next
RW vbCRLf
RW "// DONE!"
End Select
End If
RE
End Function
Function H404_session()
RTXT
Dim k,v,n,e
If Not IsObject(Session)Then
DIE "Session is disabled"
End If
e="!" ' label for "empty" value
n=Query.FQ("redirect,next")
k=TOT(Query.FQ("key,name"),rxReplace(Query.URL,"^/session[^/]*/?|/$",""))
v=Query.FQ("update,value,val,set,u,"&k)
If UBound(Split(k,"="))=1 Then
v=URLDecode(TOT(rxSelect(k,"[^\=]*$"),e))
k=rxSelect(k,"^[^\=]*")
End If
If k<>"" Then
If v<>"" Then
If v=e Then v=""
MySession.Global(k)=v
Redirect n
RW "y"'data={'"& k &"':'" & EncodeJS(v) & "'}"
Else
v=MySession.Global(k)
RW v&""
End If
Else
Select Case EXT
Case "js"
RW "data={"&vbCRLf&DIC2JS(Session.Contents)&vbCRLf&"}"
Case "xml"
RW "<session>"&vbCRLf&DIC2XML(Session.Contents)&vbCRLf&"</session>"
Case Else
RW "// LIST of session in use by '"&SERVER_NAME&"'"&vbCRLf
RW vbCRLf
For Each c In Session.Contents
RW c&"="&MySession(c)&vbCRLf
Next
RW vbCRLf
RW "// DONE!"
End Select
End If
RE
End Function
Function H404_Go()
If Query.EXT<>"" Then Exit Function
j("go-nxt")=TOT(j("REDIRECT"),FULLREQUEST)
If j("go-cmd")<>"" Then
j("go-nxt")=rxReplace(j("go-nxt"),"^/go/(("&j("go-cmd")&")/)?","/")
Else
j("go-nxt")=rxReplace(j("go-nxt"),"^/go/","/")
End If
Redirect "http://"&MyDomain&j("go-nxt")
End Function
Function H404_Go_AUTO()
MyCookies.Global("Wants.Mobile")=""
j("go-cmd")="auto"
End Function
Function H404_Go_D()
MyCookies.Global("Wants.Mobile")="n"
j("go-cmd")="d|www|desk|full|desktop"
End Function
Function H404_Go_WWW():Call H404_Go_D():End Function
Function H404_Go_FULL():Call H404_Go_D():End Function
Function H404_Go_DESK():Call H404_Go_D():End Function
Function H404_Go_DESKTOP():Call H404_Go_D():End Function
Function H404_Go_M()
MyCookies.Global("Wants.Mobile")="y"
j("go-cmd")="m|mob|mobi|lite|mobile"
End Function
Function H404_Go_MOB():Call H404_Go_M():End Function
Function H404_Go_MOBI():Call H404_Go_M():End Function
Function H404_Go_LITE():Call H404_Go_M():End Function
Function H404_Go_MOBILE():Call H404_Go_M():End Function
Function H404_debug()
MyCookies.Global("IsStaff.Flag")="y"
fwxStatus DBG.Stamp&"- fwxStatus test","hi there, this is fwxStatus test"
fwxError DBG.Stamp&"- fwxError test","hi there, this is fwxError test"
DBG.Status DBG.Stamp&"- DBG.Status"
DBG.Warning DBG.Stamp&"- DBG.Warning"
DBG.Alert DBG.Stamp&"- DBG.Alert"
DBG.Error DBG.Stamp&"- DBG.Error"
DBG.EmailStatus DBG.Stamp&"- DBG.EmailStatus subject","DBG.EmailStatus body"
DBG.EmailError DBG.Stamp&"- DBG.EmailError subject","DBG.EmailError body"
Redirect j("REDIRECT")
DBG.FatalError "FWX_VERSION="&TOT(FWX_VERSION,"(non-compiled, dev mode)")
DIE "FWX_VERSION="&TOT(FWX_VERSION,"(non-compiled, dev mode)")
RE
End Function
Function H404_robots_txt()
RTXT
FWX_ResponseStatus 200
RW "User-agent: *"&vbCRLf
RW "Sitemap: http://"&SERVER_NAME&"/sitemap.xml"&vbCRLf
For Each ptn In Split(FWX.Setting("fe.reserved.paths"),"|")
If ptn="do" Then
RW "Disallow: /"&ptn&"/"&vbCRLf
Else
RW "Disallow: /"&ptn&vbCRLf
End If
Next
RW "Disallow: /contact"&vbCRLf
For Each ptn In Split("search|share|tags|bookmark|contact|favourite|subscribe|book|feeds|popup","|")
RW "Disallow: /"&ptn&"/"&vbCRLf
Next
RW "Disallow: *RSPaging"&vbCRLf
RW "Disallow: /index.asp"&vbCRLf
RW "Disallow: /index.html"&vbCRLf
RW "Disallow: /index.htm"&vbCRLf
RW "Disallow: /index.php"&vbCRLf
For Each ptn In Split("doc|tsv|xls|xsl|rss|atom|json|asp","|")
RW "Disallow: *."&ptn&"$"&vbCRLf
Next
RW vbCRLf
RW "User-agent: AdsBot-Google"&vbCRLf
RW "Sitemap: http://"&SERVER_NAME&"/sitemap.xml"&vbCRLf
For Each ptn In Split(FWX.Setting("fe.reserved.paths"),"|")
If ptn="do" Then
RW "Disallow: /"&ptn&"/"&vbCRLf
Else
RW "Disallow: /"&ptn&vbCRLf
End If
Next
RE
End Function
Function H404():H404=True
If(Not Len(Query.URL&"")>0)Then
DBG.FatalError "Cannot resolve URL"
With Response
.Status="404 Not Found"
.ContentType="text/plain"
.Write "1. Unable to resolve request URL -"&vbCRLf
.Write "Url: "&Query.URL&vbCRLf
.Write "Query: "&Query.QUERY&vbCRLf
.Write "Raw Url: "&Query.RAWURL&vbCRLf
.End()
End With
End If
If rxTest(sUrl,"/(x|admin|bo)/")Then
DBG.FatalError "Should have gone BO"
End If
EXE "MY_H404"
Dim sUrl:sUrl=Query.URL
If rxTest(sUrl,"^/(noexist.+|google[\w\d]+)(/|\.html)$")Then
For Each key In Split("google4302d395ff1cd24e.html,"&GOOGLE_SITEMAPS_FILES&","&FWX.Setting("google.maps.key")&"",",")
If key<>"" And sUrl="/"&key Then
RHTM
FWX_ResponseStatus "200 OK"
RW "google-site-verification: "&key
RE
End If
Next
FWX_ResponseStatus "404 Not Found"
RW "google-site-verification missing page"
RE
End If
If InStr(1,sUrl,"/;")>0 Then
EXE "H404__Commands"
End If
H404__Call_Backward rxReplace(Query.URL,"^\W+|\W+$",""),""
If EXT<>"" Then
EXE "H404__"&EXT&""
End If
If IsAjax Or IsPost Or IsHuman Then
Else
If FWX.Setting("UseHTM")<>"" Then
If Query.EXT="" Then
sUrl=Replace(sUrl&"/","//","/")
sUrl=Replace(sUrl&".*tm","/.*tm",".htm")
PermanentRedirect sUrl&IIf(sQry<>"","?"&sQry,"")
End If
Else
If rxTest(sUrl,"\.html?$")Then
sUrl=rxReplace(sUrl,"\.html?$","/")
PermanentRedirect sUrl&IIf(sQry<>"","?"&sQry,"")
End If
End If
End If
End Function
Function H404__Call(ByVal s)
Dim b:b=""
s=rxReplace(Trim(s),"[^\w\_\(\)\""]+","_")
If s<>"" Then
b=EXE("MY_H404_"&s&"")
If b="" Then
b=EXE("H404_"&s&"")
End If
End If
H404__Call=b
End Function
Function H404__Call_Backward(ByVal s,ByVal f)
If f<>"" Then f="_"&f
Dim o
s=rxReplace(Trim(s),"[^\w\_]+","_")
Do While s<>""
o=TOT(o,H404__Call(s&f))
If InStrRev(s,"_")>1 Then
s=Left(s,InStrRev(s,"_")-1)
Else
s=""
End If
Loop
H404__Call_Backward=o
End Function
Dim Query:Set Query=New fwxRequest
With Query
.Initialize()
Dim RAWURL:RAWURL=.RAWURL
Dim RAWQRY:RAWQRY=.RAWQRY
Dim RAWURI:RAWURI=.RAWURI
Dim URI:URI=.URI
Dim URL:URL=.URL
Dim EXT:EXT=.EXT
Dim FULLURL:FULLURL=.FULLURL
Dim PAGE:PAGE=.PAGE
Dim PAGEID:PAGEID=.PAGEID
Dim PATH:PATH=.PATH
Dim PATHID:PATHID=.PATHID
Dim FULLID:FULLID=.FULLID
End With
Function xIFC_Config():xIFC_Config=True
If IsPost Then
Application.Contents.RemoveAll()
FWX.ServerSettings=Query.F("Server")
FWX.DomainSettings=Query.F("Domain")
FWX.Reboot()
If IsAjax Then RC:DIE "y":End If
End If
rw "<scr"&"ipt type='text/javascript'>"
rw "if(jQuery) (function($){$(function(){$('#reset').click(function(){top.location='?|do=reset';});$('#save').click(function(){$(this.form).ajaxSubmit({success:function(s){fwx.Y('Settings saved!', {timer:3});}});return false;});});})(jQuery);"
rw "</scr"&"ipt>"
rw "<form action='"&SCRIPT&"?|do=Config' method='POST'>"
rw "<table width='100%' cellspacing='0'>"
rw "<tr>"
rw "<td valign='top' width=''>"
rw "<div id='config-tab-Server'>"
rw "<div class='Info P3 B'>LEVEL 1: Server Settings</div>"
rw "<div align='left'>"
rw "<textarea wrap='off' name='Server' id='Server' style='padding:0;margin:0;border:#ccc solid 5px;width:98%;height:100px;'>"&Encode(FWX.ServerSettings)&"</textarea>"
rw "</div>"
rw "</div>"
rw vsp(20)
rw "<div id='config-tab-Instance'>"
rw "<div class='Info P3 B'>LEVEL 2: Instance Settings</div>"
rw "<div align='left'>"
rw "<div style='padding:0;margin:0;border:#ccc solid 5px;width:98%;' class='P3 Disabled'>...loaded from central database</div>"
rw "</div>"
rw "</div>"
rw vsp(20)
rw "<div id='config-tab-Domain'>"
rw "<div class='Info P3 B'>LEVEL 3: Domain Settings</div>"
rw "<div align='left'>"
rw "<textarea wrap='off' name='Domain' id='Domain' style='padding:0;margin:0;border:#ccc solid 5px;width:98%;height:120px;'>"&Encode(FWX.DomainSettings)&"</textarea>"
rw "</div>"
rw "</div>"
rw "<div class='TRight P5 RowC zB5'>"
rw "<input id='reset' value='Reboot system' type='button' style='color:#009;'/>"
rw "<input id='save' value='Apply changes &raquo;' type='submit' style='color:#090;' class='B'/>"
rw "</div>"
rw "</td>"
rw "<td valign='top' width='20'>"& hsp(20)&"</td>"
rw "<td valign='top' width='200'>"
rw "<div id='config-tab-Security'>"
rw "<div class='Info P3 B'>Security Information:</div>"
rw "<div align='left'>"
rw "<a href='javascript:;' onclick='window.prompt(this.innerHTML,this.title)' title='"&FWX.Serial&"'>Serial Number</a>"
rw " | "
rw "<a href='javascript:;' onclick='window.prompt(this.innerHTML,this.title)' title='"&FWX.DomainKey&"'>Domain ID</a>"
rw "</div>"
rw "</div>"
rw vsp(20)
rw "<div id='config-tab-Current'>"
rw "<div class='Info P3 B'>Running Configuration:</div>"
rw "<div align='left'>"
Dim s
For Each s In Application.Contents
If Left(s,1)<>"-" Or Left(s,Len(FWX.DomainPrefix()))=FWX.DomainPrefix()Then
rw "<div class='NoBR' style='border-bottom:dotted 1px;'>"
rw "<a href='javascript:;'"
rw " onclick=""window.prompt('"&EncodeJS(s)&"',this.title)"""
rw " title="""&EncodeJS(Encode(Application.Contents(s)))&""""
rw ">"&s&"</a> ("&Len(Application.Contents(s))&"b)"
rw "</div>"
End If
Next
rw "</div>"
rw "</div>"
rw vsp(10)
rw "</td>"
rw "</tr>"
rw "</table>"
rw "</form>"
End Function
Function xIFC_START():xIFC_START=True
rw "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
rw "<html lang=""en"" xml:lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"">"
rw "<html>"
rw "<head>"
rw "<link type='text/css' rel='stylesheet' media='all' href='/fwx/@.css'/>"
rw "<style>"
rw "/* LAYOUT - Layer 0:sticky footer */"
rw "html,body{ height:100%; }"
rw "#bo-head{ position:fixed; width:100%; top:0; }"
rw "#bo-foot{ position:fixed; width:100%; bottom:0; }"
rw "html > body #bo-body{ padding:35px 0 35px 0; }"
rw "/* SIZES */"
rw ".bo-head{ width:980px; margin:0 auto; padding:3px 0; }"
rw ".bo-body{ width:980px; margin:0 auto; }"
rw ".bo-foot{ width:980px; margin:0 auto; padding:3px 0; }"
rw "/* COLOURS/STYLE */"
rw "#bo-head{ background:#f7f7f7; border-bottom:#ccc solid 3px; z-index:10; }"
rw "#bo-body{ border-color:#ccc; }"
rw "#bo-foot{ background:#f7f7f7; border-top:#ccc solid 3px; z-index:10; }"
rw "</style>"
rw "<scr"&"ipt type='text/javascript' src='/fwx/@.js'></scr"&"ipt>"
rw "</head>"
rw "<body style='padding:0;margin:0;'>"
rw "<div class='Clear' id='bo-head'>"
rw "<div class='Clear bo-head'>"
rw "<span class='R'>"
rw "<a href='"&SCRIPT&"?|do=logout'>Logout</a>"
rw "</span>"
rw "<a href='"&SCRIPT&"?|do=config'><strong>Fyneworks Setup Utility</strong></a>"
rw "</div>"
rw "</div>"
rw "<div class='Clear' id='bo-body'>"
rw "<div class='Clear bo-body'>"
End Function
Function xIFC_END():xIFC_END=True
rw "</div>"
rw "</div>"
rw "<div class='Clear' id='bo-foot'>"
rw "<div class='Clear bo-foot'>"
rw "<span class='R'>"
rw "(c) 2003-"&Right(Year(Now()),2)&" Fyneworks <a href='http://www.fyneworks.com/'>CMS and SEO</a>"
rw "</span>"
rw "Domain: <a href='http://"&SERVER_NAME&"/'>"&MyDomain&"</a>"
rw "</div>"
rw "</div>"
rw "</body>"
rw "</html>"
End Function
Function xIFC_Reset():xIFC_Reset=True
Application.Contents.RemoveAll()
Session.Abandon()
rw "<strong>IIS Application Reset</strong><br/>"
rw "<a href='?|do=Login'>Config &raquo;</a><br/>"
rw "<a href='/admin/'>Admin &raquo;</a><br/>"
rw "<a href='/'>Home</a><br/>"
End Function
Function xIFC_Login():xIFC_Login=True
rw "<p>"
rw "<strong class='No'>This is a strictly retricted area</strong>"
rw "</p>"
rw ""
rw "<form action='"&SCRIPT&"?|do=login' method='POST'>"
rw "<input name='redirect' value='"&Replace(FULLREQUEST,"|do=login","|do=config")&"' type='hidden'/>"
rw "<table width='' cellspacing='10'>"
rw "<tr>"
rw "<td align='right'><label for='xdomain'>Domain:</label></td>"
rw "<td><a href='http://"&SERVER_NAME&"/'><strong>"&MyDomain&"</strong></a></td>"
rw "</tr>"
rw "<tr>"
rw "<td align='right'><label for='xauth'>Password:</label></td>"
rw "<td>"
rw "<input id='xauth' name='xauth' value='"&Query.FQ("auth")&"' type='password'/>"
If IsLocal Or FWX.Setting("password")="" Then
rw " (<a href='javascript:;' onclick=""(function(s){ if(s) $('#xauth').val(s); })(window.prompt(this.innerHTML,this.title))"" title='"&FWX.DomainKey&"'>Domain ID</a>)"
End If
rw "</td>"
rw "</tr>"
rw "<tr>"
rw "<td align='right'>&nbsp;</td>"
rw "<td><input value='Login &raquo;' type='submit' class='Yes'/></td>"
rw "</tr>"
rw "</table>"
rw "</form>"
End Function
Function xIFC_jsessionid():xIFC_jsessionid=True
PermanentRedirect "/"
End Function
Class c_xIFC
Public Function Go(ByVal ACTION)
Dim ACTION_URL:ACTION_URL=SCRIPT&"?|do="&ACTION&""
Response.Write "<a href='"&ACTION_URL&"'><strong>FWX Action &raquo;</strong></a>"
If IsLive Then
Response.Write "<scr"&"ipt>window.location='"&ACTION_URL&"'</scr"&"ipt>"
End If
Response.End()
End Function
Public Function Authorized(ByVal PASSWORD)
Authorized=False
If FWX.Setting("password")<>"" Then
If PASSWORD=FWX.Setting("password")Then Authorized=True
Else
If OneWayEncrypt(PASSWORD)="191f42bcb3f8095f8969c5cf5824a01488a36fba939d414320f0cc2b7307a64a" Then
PASSWORD=FWX.DomainKey
End If
If PASSWORD=FWX.DomainKey Then
Authorized=True
Else
If PASSWORD="fyneworks" Then
WriteToFile "/@OTA("&FWX.DomainKey&").asp",BeginASP&" FWX_AddHeader ""Status"",""301"" : FWX_AddHeader ""Location"",""/"""&EndASP
End If
End If
End If
End Function
Private Sub Class_Initialize()
Dim d
d=LCase(Query.FQ("|do"))
If d<>"" Then
Response.ExpiresAbsolute=Date()-1
If d="logout" Then
Session.Abandon()
Go "login"
End If
Dim AUTH
AUTH=TOT(Query.FQ("xauth"),Decrypt(MySession("xauth")))
If Me.Authorized(AUTH)Then
MySession("xauth")=Encrypt(AUTH)
If d="login" Then
Redirect j("REDIRECT")
Go "config"
End If
Else
If d="reset" Then
Else
Response.ContentType="text/html"
d="login"
End If
End If
d=rxReplace(d,"[^\w\_]+","_")
EXE "xIFC__"&d&""
If Not EXE("xIFC_"&d&"_"&"START")Then EXE "xIFC_"&"START"
EXE "xIFC_"&d&""
If Not EXE("xIFC_"&d&"_"&"END")Then EXE "xIFC_"&"END"
RE
End If
End Sub
End Class
Dim xIFC
Set xIFC=New c_xIFC
xIFC=""
For Each Var In Split("PAGE,PAGEID,PATH,PATHID,FULLID,FULLREQUEST,SCRIPT,URL,RAWURL,FULLURL,Is404,IsAjax,IsApp,IsFront,IsBack,IsDesktop,IsMobile,IsLocal,IsLive,IsTest,Query.QString,Query.FString,MySite,MyDomain,MyDomainID,MyWWWDomain,FIXD_SERVER,THIS_SERVER,FIXD_DOMAIN,THIS_DOMAIN,DOMAIN_NAME,SERVER_NAME,IsSubDomain,IsBOSub,SUBDOMAIN,CLIENT_IP,SERVER_IP,USERAGENT,UAID,OSID,REFERRER,IsHuman,IsRobot,IsMobile,IsDesktop,MyAccount,MyEmail,MyUsername,MyName,TMPL_FOLDER,SITE_FOLDER,",",")
If Var<>"" Then
j(Var)=EXE(Var)
If j(Var)<>"" And IsNumeric(j(Var))Then j(Var)=VOZ(j(Var))
End If
Next
j("FDATA")=Query.FString
j("QDATA")=Query.QString
j("FString")=Query.FString
j("QString")=Query.QString
j("SERVER")=SERVER_NAME
j("DOMAIN")=DOMAIN_NAME
j("EDITION")=rxReplace(FWX.SETTING("EDITION"),"[^0-9a-zA-Z]+","")
j("REVISION")=rxReplace(FWX.SETTING("REVISION"),"[^0-9a-zA-Z]+","")
j("REV")=rxReplace("E"&j("EDITION")&"R"&j("REVISION"),"[^0-9a-zA-Z]+","")
j("BOOT-TIME")=FWX.SETTING("BOOT-TIME")
IsRaw=CBool(IsRaw Or Query.FQ("-raw")<>""):j("IsRaw")=IsRaw
IsMini=CBool(IsMini Or Query.FQ("-mini")<>""):j("IsMini")=IsMini
IsApp=CBool(IsApp Or Query.FQ("-ajax")<>""):j("IsApp")=IsApp
IsAjax=CBool(IsApp Or IsAjax Or Query.FQ("-ajax")<>""):j("IsAjax")=IsAjax
If IsAjax Then
j("DO-NOT-CACHE")="y"
Query.Data("-ajax")="y"
End If
j("REDIRECT")=Query.FQ("redirect,next,continue,redir,r")
If VOZ(j("REDIRECT"))<>0 Or j("REDIRECT")="0" Then j("REDIRECT")=""
CONST TAX_RATE=20.0
If j("TAX_RATE")="" Then
j("TAX_RATE")=TOT(TOT(FWX.Config("vartax"),j("TAX_RATE")),TAX_RATE)
End If
Dim OSID:OSID=OSTag(USERAGENT)
Dim UAID:UAID=UATag(USERAGENT)
If Is404 Then
SCRIPT=URL
If EXT="htm" Or EXT="" Then
SCRIPT=rxReplace(SCRIPT,"(/|\.html?)$",".asp")
End If
Call H404()
End If
Function DEFAULT_DOMAIN_NAME_ENFORCEMENT()
Dim x,y,ForceNoWWW,FavorNoWWW
x=SERVER_NAME
y=x
ForceNoWWW=IsTrue(FWX.Setting("NoWWW"))
FavorNoWWW=IsTrue(CBool(IsTrue(IsBOSub)Or CBool(InStr(1,MyDomainTLDLess,".")>0)))
If IsLocal Or StartsWith(y,"local.")Or EndsWith(y,".local")Or y="localhost" Then
Exit Function
ElseIf IsInList(y,FWX.Setting("domain.alt"))And y<>MyDomain Then
Exit Function
ElseIf IsInList(y,FWX.Setting("domain.alias"))And y<>MyDomain Then
Exit Function
ElseIf rxTest(y,"azurewebsites")Then
Exit Function
ElseIf Left(y,4)="www." Then
If(IsSubDomain And(IsCDN Or IsTest Or IsMobile))Or ForceNoWWW Then
y=Mid(y,5)
Else
If FavorNoWWW Then
y=Mid(y,5)
Else
End If
End If
Else
If 1=0 Then
ElseIf FavorNoWWW Then
ElseIf IsTest Then
If IsRobot And Not(IsTestBot Or rxTest(USERAGENT,"(testomato|Screaming Frog)"))Then
y="www."&MyDomain
End If
ElseIf IsMobile Then
ElseIf IsCDN Then
If Query.EXT="" Or rxTest(Query.EXT,"asp|html?|xml|rss|doc|pdf")Then
y="www."&rxReplace(y,"^(w\d)\.","")
If ForceNoWWW Then y=Mid(y,5)
Else
End If
Else
y="www."&y&""
End If
End If
If IsHuman And(IsTemp)Then
Else
If IsLocal Then
ElseIf IsBoSub Then
ElseIf IsInList(y,FWX.Setting("domain.alias"))Then
ElseIf DOMAIN_NAME=MyDomain Then
Else
If SERVER_NAME="localhost" Then
Else
y=rxReplace(y,"\b"&rxEscape(DOMAIN_NAME)&"$",MyDomain)
If y=MyDomain Then y=MyWWWDomain
If x=y Then
If ForceNoWWW Then
y=MyDomain
Else
If SERVER_NAME="localhost" Then
Else
y="www."&MyDomain
End If
End If
Else
End If
End If
End If
End If
If y<>x And y<>"" Then
Dim z
z=FWX.Config("domain.alt")
If z="" Then z=Application.Contents("domain.alt")
If z="" Then z=FWX__domain_alt
If z<>"" And x=z Then
y=x
End If
End If
If y<>x And y<>"" Then
y=TOT(PROTOCOL,"http")&"://"&y&Query.URL&IIf(Query.QUERY<>"","?"&Query.QUERY,"")
If IsHTTPReq Then
Else
PermanentRedirect y
End If
End If
End Function
Function DEFAULT_SUB_DOMAIN_NAME_BEHAVIOUR()
If rxTest(SERVER_NAME,"^((l\d|local|test|temp|dev)\.)?("&FWX__BO_Reserved_SubDomains&")\.")Then
DBG "This is a Reserved SubDomains"
If FWX__Ignore_SubDomain_Behaviour Then
DBG "IGNORING SUB-DOMAI BEHAVIOUR ENFORCEMENT"
Exit Function
End If
If IsHuman Then
If rxTest(QUERY_URL,"^/(x|admin)")Then
If rxTest(SERVER_NAME,"\b(driver|drivers|dp)\b")Then j("@ENGINE")="DriversPortal"
ElseIf 1=0 Or rxTest(QUERY_URL,"^/(basket|cart|my|users?|clients?|checkout|pay)\b")Or rxTest(QUERY_URL,"\b(callback|ipn)\b")Or rxTest(QUERY_URL,"\b(quote|invoice)\b")Or rxTest(QUERY_URL,"\.(ico)$")Then
If IsStaffPC Then
Else
Redirect MyProtocol&"://www."&rxReplace(SERVER_NAME,"\b("&FWX__BO_Reserved_SubDomains&")\.","")&FULLREQUEST
End If
Else
If REFERRER<>"" And IsStaffPC Then
Else
If rxTest(SERVER_NAME,"cms\.")Then
ElseIf Not Left(QUERY_URL,3)="/x/" Then
Redirect "/x"&QUERY_URL
Else
End If
End If
End If
Else
PermanentRedirect MyProtocol&"://www."&rxReplace(SERVER_NAME,"\b("&FWX__BO_Reserved_SubDomains&")\.","")&FULLREQUEST
End If
End If
End Function
DBG "DOMAIN_NAME_ENFORCEMENT > GLOBAL_DOMAIN_DISPATCHER > running from "&SCRIPT_NAME
If GLOBAL_DOMAIN_DISPATCHER Then
DBG "DOMAIN_NAME_ENFORCEMENT > GLOBAL_DOMAIN_DISPATCHER > CUSTOM OVERRIDE APPLIED!"
Else
DBG "DOMAIN_NAME_ENFORCEMENT > GLOBAL_DOMAIN_DISPATCHER > CALL DEFAULT_GLOBAL_DOMAIN_DISPATCHER"
End If
DBG "DOMAIN_NAME_ENFORCEMENT > SUB_DOMAIN_NAME_BEHAVIOUR > running from "&SCRIPT_NAME
If SUB_DOMAIN_NAME_BEHAVIOUR Then
DBG "DOMAIN_NAME_ENFORCEMENT > SUB_DOMAIN_NAME_BEHAVIOUR > CUSTOM OVERRIDE APPLIED!"
Else
DBG "DOMAIN_NAME_ENFORCEMENT > SUB_DOMAIN_NAME_BEHAVIOUR > CALL DEFAULT_SUB_DOMAIN_NAME_BEHAVIOUR"
Call DEFAULT_SUB_DOMAIN_NAME_BEHAVIOUR()
End If
If IsHTTPReq Then
DBG "DOMAIN_NAME_ENFORCEMENT = SKIPPING because IsHTTPReq="&IsHTTPReq
Else
DBG "DOMAIN_NAME_ENFORCEMENT = running from "&request.ServerVariables("SCRIPT_NAME")
If SHARED_DOMAIN_ENFORCEMENT=True Then
DBG "SHARED_DOMAIN_ENFORCEMENT = SHARED_DOMAIN_ENFORCEMENT APPLIED!"
ElseIf DOMAIN_NAME_ENFORCEMENT=True Then
DBG "DOMAIN_NAME_ENFORCEMENT = CUSTOM OVERRIDE APPLIED!"
Else
DBG "DOMAIN_NAME_ENFORCEMENT = CALL DEFAULT_DOMAIN_NAME_ENFORCEMENT"
Call DEFAULT_DOMAIN_NAME_ENFORCEMENT()
End If
End If
If Query.FQ("--stockist")<>"" Then CACHE__quoteStockist=Query.FQ("--stockist")
If Query.FQ("--merchant")<>"" Then CACHE__quoteMerchant=Query.FQ("--merchant")
EXE "ENVIRONMENT_READY"
EXE "PROCESS_REQUEST"%>