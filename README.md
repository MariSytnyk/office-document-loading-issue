# document-loading-issue

1. Install the dependencies in the local node_modules folder.

        npm install

2. Start dev server by running the command
        
        npm run dev-server

3. Sideload add-in by running the command

        npm run start

   or by uploading manifest file to web version of Word.

4. To see the issue choose the document with several pages and with pictures in th end of document. 
5. Open add-in, open the console and press Run button.  

###As a result you should see this error:
```
Error:  RichApi.Error
   at new n (word-web-16.00.js:25:309056)
   at i.processRequestExecutorResponseMessage (word-web-16.00.js:25:373222)
   at word-web-16.00.js:25:371285
```
