
var strs=[];

function vbsTojs(vbs){

	
    var s = vbs;
    var Vars = '';
    var Fx = '';
    var FxHead = '';
    var Args = '';


//prep function block
    s = s.replace(/^sub/gim, "function");
    s = HideStrings(s);
    //s = s.match(/(?:Function|Sub)[\w\W]+End\s+(?:Function|Sub)/gim)[0];
    s = s.replace(/\&/gm, "+");
    s = s.replace(/_\n/gm,"");
    s = s.replace(/:/gm,"\n");
    //single line IF statements need to go to multiple lines
    s = s.replace(/\bthen\b[ \t](.+)/gi,"then\n$1\nEnd If");
    
    //split block into separate lines
    a=s.split('\n');

    //trim spaces and remove empty lines
    for(i=0;i<a.length;i++){
        a[i]=a[i].replace(/^\s+|\s+$/,"");}
    a = a.filter(function (val) { return val !== ""; });


//Fix FUNCTION tags
    a[0]=a[0].replace(/function\s+/i,"");
    Fx = a[0].match(/^\w+/)[0];
    a[0]=a[0].replace(Fx,"").replace(/[\(\)]/g,"");
    a[0]=a[0].replace(/\bbyval\b/gi,"").replace(/\bbyref\b/gi,"").replace(/\boptional\b/gi,"");
    a[0]=a[0].replace(/\bas\s+\w+\b/gi,"");
    a[0]=a[0].replace(/\s+/g,"");
    a[0]=a[0].replace(/,/gi,", ");
    FxHead = "function " + Fx+ " ("+ a[0] + "){";
    a[0]="";
    //Remove END FUNCTION tags
    a.length = a.length-1;
	

//Fix Syntax
    for(i=1;i<a.length;i++){
	
        //Vars
        if(a[i].search(/^dim\s+/i)>-1){
            a[i]=a[i].replace(/dim\s*/i,"");
            Vars += a[i] + ",";
            a[i]='';

            //MSGBOX
        } else if (a[i].search(/^msgbox\(/i) > -1) {
            a[i] = a[i].replace(/^msgbox\(/i, "alert(");

        } else if (a[i].search(/(^msgbox\s)+(.*)/i) > -1) {
            a[i] = a[i].replace(/(^msgbox\s)+(.*)/i, "alert($2)");


            //INPUTBOX
        } else if (a[i].search(/^inputbox\(/i) > -1) {
            a[i] = a[i].replace(/^inputbox\(/i, "prompt(");

        } else if (a[i].search(/(^inputbox\s)+(.*)/i) > -1) {
            a[i] = a[i].replace(/(^inputbox\s)+(.*)/i, "prompt($2)");


       //SUB AND CALL
        } else if (a[i].search(/^sub\s/i) > -1) {
            a[i] = a[i].replace(/^sub\s/i, "");
        } else if (a[i].search(/\bEnd Sub\b/i) > -1) {
            a[i] = a[i].replace(/\bEnd Sub\b/i, "}");
        } else if (a[i].search(/\bcall \b/i) > -1) {
            a[i] = a[i].replace(/\bcall \b/i, "");
            
        //OBJECTS
        } else if (a[i].search(/\\sSET\\s+/i) > -1) {
            a[i] = a[i].replace(/\\sSET\\s+/i, "");
        } else if (a[i].search(/'nothing'/i) > -1) {
            a[i] = a[i].replace(/'nothing'/i, "null");

            //FOR
        }else if(a[i].search(/^\bFOR\b\s+/i)>-1){
            a[i]=a[i].replace(/^\bFOR\b\s+/i,"");
            counter = a[i].match(/^\w+/)[0];
            from = a[i].match(/=\s*[\w\(\)]+/)[0];
            from=from.replace(/=/,"").replace(/\s+/g,"");
            a[i]=a[i].replace(counter,"").replace(from,"").replace(/\bTO\b/i,"");
            to = a[i].match(/\s*[\w\(\)]+\s*/)[0];
            to=to.replace(/=/,"").replace(/\s+/g,"");
            // stepsize
            if(a[i].search(/\bSTEP\b\s+/i)==-1) a[i] = a[i] + " STEP 1";
            steps = a[i].match(/\bSTEP\b\s*(-?)([*+-\/\s\w\(\)]*)$/i);
            steps[1] = steps[1].replace(/\s*/g,"")=="" ? "+" : "-";
            cmpop = (steps[1] == "+") ? "<=" : ">=";
            if (steps[2] == "1") {
              stepsize = counter + steps[1] + steps[1];
            }else{
              stepsize = (counter + " = " + counter + " " + steps[1] + " " + steps[2]).replace(/\s*/g," ");
            }
            
            a[i] = "for(" + counter + "=" + from + "; " + counter + cmpop + to + "; " + stepsize + "){";
	
            //NEXT
        }else if(a[i].search(/^NEXT\b/i)>-1){
            a[i] = "}";


            //EXIT FOR
        }else if(a[i].search(/\bEXIT\b\s*\bFOR\b/i)>-1){
            a[i] = "break";
	
            //IF
        } else if (a[i].search(/(IF\\s+\\S+\\s*)=(\\s*\\S+\\s+THEN)/i) > -1) {
            a[i] = a[i].replace(/(IF\\s+\\S+\\s*)=(\\s*\\S+\\s+THEN)/i, "$1==$2");
        }else if(a[i].search(/^\bIF\b\s+/i)>-1){
            a[i]=a[i].replace(/^\bIF\b\s+/i,"");
            a[i]=a[i].replace(/\bTHEN$\b/i,"");
            a[i]=a[i].replace(/=/g,"==").replace(/<>/g,"!=");                 
            a[i]=a[i].replace(/\bOR\b/gi,"||").replace(/\bAND\b/gi,"&&");     
            a[i] = "if(" + a[i] + "){";
	
            //ELSE and ELSEIF
        }else if(a[i].search(/^ELSE/i)>-1){
            a[i] = a[i].replace(/^ELSE/i, "}else{");
            a[i] = a[i].replace(/=/g, "==").replace(/<>/g, "!=");
            a[i] = a[i].replace(/\bOR\b/gi, "||").replace(/\bAND\b/gi, "&&");
        } else if (a[i].search(/^ELSEIF/i) > -1) {
            a[i] = a[i].replace(/^ELSEIF/i, "}else if{");
            a[i] = a[i].replace(/=/g, "==").replace(/<>/g, "!=");
            a[i] = a[i].replace(/\bOR\b/gi, "||").replace(/\bAND\b/gi, "&&");

            //END IF
        }else if(a[i].search(/^END\s*IF/i)>-1){
            a[i] = "}";

            //WHILE
        }else if(a[i].search(/^WHILE\s/i)>-1){
            a[i] = a[i].replace(/^WHILE(.+)/i, "while($1){");

            //WEND
        }else if(a[i].search(/^WEND/i)>-1){
            a[i] = "}";

            //DO WHILE
        }else if(a[i].search(/^DO\s+WHILE\s/i)>-1){
            a[i] = a[i].replace(/^DO\s+WHILE(.+)/i, "while($1){");

            // there is no DO UNTIL equiv in JavaScript - use DO WHILE instead
        } else if (a[i].search(/^DO\s+UNTIL\s/i) > -1) {
            a[i] = a[i].replace(/^DO\s+WHILE(.+)/i, "while($1){");

            //LOOP
        }else if(a[i].search(/^LOOP$/i)>-1){
            a[i] = "}";

            //EXIT FUNCTION
        }else if(a[i].search(/\bEXIT\b\s*\bFUNCTION\b/i)>-1){
            a[i] = "return";

            //STEP
        } else if (a[i].search(/\sSTEP\s/i) > -1) {
            a[i] = a[i].replace(/\sSTEP\s/i, "+");

            //SELECT CASE
        }else if(a[i].search(/^SELECT\s+CASE(.+$)/i)>-1){
            a[i]=a[i].replace(/^SELECT\s+CASE(.+$)/i,"switch($1){");
        }else if(a[i].search(/^END\s+SELECT/i)>-1){
            a[i] = "}";
        }else if(a[i].search(/^CASE\s+ELSE/i)>-1){
            a[i] = "default:";
        }else if(a[i].search(/^CASE[\w\W]+$/i)>-1){
            a[i] = a[i] + ":" ;
        

            //ERROR HANDLING
        } else if (a[i].search(/^On\\s+Error\\s+Resume\\s+Next.*[\r\n]/i) > -1) {
            a[i] = a[i].replace(/^On\\s+Error\\s+Resume\\s+Next.*[\r\n]/i, "window.onerror=null\r\n");
        } else if (a[i].search(/^On\\s+Error\\s+.+.*[\r\n]/i) > -1) {
            a[i] = a[i].replace(/^On\\s+Error\\s+.+.*[\r\n]/i, "window.detachEvent('onerror')\r\n");


            //COMPARISON
        } else if (a[i].search(/(?=\\s*)&(?!#|[a-z]+;)/i) > -1) {
            a[i] = a[i].replace(/(?=\\s*)&(?!#|[a-z]+;)/i, "+");
        } else if (a[i].search(/(\s+)NOT(\s+)/i) > -1) {
            a[i] = a[i].replace(/(\s+)NOT(\s+)/i, "$1!$2");
        } else if (a[i].search(/(\s*)<>(\s*)/i) > -1) {
            a[i] = a[i].replace(/(\s*)<>(\s*)/i, "$1!=$2");
        } else if (a[i].search(/(\s+)AND(\s+)/i) > -1) {
            a[i] = a[i].replace(/(\s+)AND(\s+)/i, "$1&&$2");
        } else if (a[i].search(/(\s+)OR(\s+)/i) > -1) {
            a[i] = a[i].replace(/(\s+)OR(\s+)/i, "$1||$2");
        
        
            //OPTION EXPLICIT AND CONST
        } else if (a[i].search(/^CONST/i) > -1) {
            a[i] = a[i].replace(/^CONST/i,"const");
        } else if (a[i].search(/^Option\\s+Explicit.*[\r\n]/i) > -1) {
            a[i] = a[i].replace(/^Option\\s+Explicit.*[\r\n]/i, "");
    }		
            
        else{
            //alert(a[i]);
        }
		      

    }
	

    //OTHER STUFF
    for (i = 0; i < a.length; i++) {
        //comments
        a[i] = a[i].replace(/^\'/i, "//");
        //attempt to catch inline comments
        a[i] = a[i].replace(/\s\s\'/i, "  //");

        a[i] = a[i].replace(/\sByVal\s/i, " ");
        a[i] = a[i].replace(/\sByRef\s/i, " ");
        a[i] = a[i].replace(/vbCRLF/i, "\\r\\n");
        a[i] = a[i].replace(/vbCR/i, "\\r");
        a[i] = a[i].replace(/vbLF/i, "\\n");
        a[i] = a[i].replace(/vbTab/i, "\\t");
        a[i] = a[i].replace(/vbOK/i, "1");
        a[i] = a[i].replace(/vbCancel/i, "2");
        a[i] = a[i].replace(/vbAbort/i, "3");
        a[i] = a[i].replace(/vbRetry/i, "4");
        a[i] = a[i].replace(/vbIgnore/i, "5");
        a[i] = a[i].replace(/vbYes/i, "6");
        a[i] = a[i].replace(/vbNo/i, "7");
        a[i] = a[i].replace(/vbBinaryCompare/i, "0");
        a[i] = a[i].replace(/vbTextCompare/i, "1");
        a[i] = a[i].replace(/vbUseDefault/i, "-2");
        a[i] = a[i].replace(/vbTrue/i, "-1");
        a[i] = a[i].replace(/vbFalse/i, "0");
    }



    //alert(a.join("*"));	
	Vars = Vars.replace(/\s*AS\s+\w+\s*/gi,"");
    if(Vars!="") Vars = "var " + Vars.replace(/,$/,";").replace(/,/g,", ");
    FxHead  + '\n' + Vars;
	
    a=a.filter(function(val) { return val !== ""; }) //remove empty items
	
    for(i=0;i<a.length;i++){
        if (a[i].search(/[^}{:]$/)>-1) a[i]+=";";
    }
	
    ss = FxHead + '\n' + Vars + '\n' + a.join('\n') + '\n}';
	
    ss = ss.replace(new RegExp(Fx+"\\s*=\\s*","gi"),"return ");
	
    ss = UnHideStrings(ss);

    return jsIndenter(ss);
}




//-----------------------------------------------------

function jsIndenter(js){

    var a=js.split('\n	var');
    var s = '';

    //trim
    for(i=0;i<a.length;i++){ a[i]=a[i].replace(/^\s+|\s+$/,""); }
    //remove empty items
    a=a.filter(function(val) { return val !== ""; });


    for(var i=1;i<a.length;i++){

        if(a[i-1].indexOf("{")>-1) margin += 4 ;
       
        if(a[i].indexOf("}")>-1) { margin -= 4; }
        
        if(margin<0) margin = 0;

        a[i] = StrFill(margin," ") + a[i] ;
    }
    return a.join('\n');
}


function StrFill(Count,StrToFill){
    var objStr,idx;
    if(StrToFill=="" || Count==0){
        return "";
    }
    objStr="";
    for(idx=1;idx<=Count;idx++){
        objStr += StrToFill;
    }
    return objStr;
}

function HideStrings(text){

    const x = String.fromCharCode(7);
    const xxx = String.fromCharCode(8);

    text = text.replace(/"""/gim, '"'+xxx);  //hide 3 quotes " " "
    var idx=0, f=0;
    while(f>-1){
        f = text.search(/".+?"/gim);
        if(f>-1){
            strs.push(text.match(/".+?"/)[0]);
            //alert(strs[idx]);
            text = text.replace(/".+?"/, x+idx+x);
            idx++;
        }
    }
    //alert(text);
    return text;
}

function UnHideStrings(text){
    for(var i=0; i<strs.length; i++){
        text = text.replace(new RegExp("\\x07"+i+"\\x07"), strs[i]);
    }
    //Unhide 3 quotes " " " ***BUG: causes unterminated string if triple-quotes are at the end of the string
    text = text.replace(/\x08/gim,'\\"');
    text = text.replace(/""/gi,'\\"');    
    return text;
}

