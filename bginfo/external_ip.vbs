ON ERROR RESUME NEXT
CONST conURLSrc = "http://checkip.dyndns.org/"
SET http = createObject("microsoft.xmlhttp") 
SET fSO = createObject("Scripting.FileSystemObject") 

http.open "GET",conURLSrc,FALSE 
http.send 
htmlIP = http.responseText
SET regEx = NEW regExp 
regEx.global = TRUE 
regEx.pattern =  "Current IP Address: ([0-9\.]*)"
result = regEx.test(htmlIP)
SET matches = regEx.execute(htmlIP) 
FOR EACH match IN matches
        extIP = match.subMatches(0)
NEXT
ECHO extIP
