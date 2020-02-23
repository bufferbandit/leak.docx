# leak.docx
Leak windows system info through a docx file.

You can get info like the architecture (through the "ua-cpu" header), office version (through the "x-office-major-version" header) and ofcourse the IP. I've seen some examples of getting the Windows version through the "user-agent" header, but these requests dont add that. My hypothesis is that Word =< 14 used to send these headers but I haven't tested this.

Funny side effect is that the text in the docx file will be updated to the HTTP response
