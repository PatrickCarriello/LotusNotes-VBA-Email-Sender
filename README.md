# LotusNotes-VBA-Email-Sender
A collection of VBA functions for sending email through IBM Lotus Notes.

When looking for more complete functions for sending e-mail through Lotus Notes, i realized that there is a lot of information mismatched and loose on the internet.
I found many pieces, but nothing too complex. So I decided to program my own module with functions that not only meet me, but also meet those who seek them.

Feel free to use it in the best way possible and contribute to what you discover new. You can also enhance the functions present here.

<h2>Tests</h2>

All tests were done (and worked) in the following versions:

<uh>
<li>Microsoft Office 2013</li>
<li>IBM Lotus Notes (now just a IBM Notes) 9</li>
<li>Microsoft XML, v6.0</li>
</uh>
<br>
Probably the functions also work in other versions or just need a few changes.

<h2>Using Functions</h2>

<uh>
<li><strong>SendEmail(subject As String, body As String, emails() As Variant)</strong></li>

<pre><code>Dim emailsendto() as String
Dim counter as Integer
counter = 2
ReDim emailsendto(2)
emailsendto(0) = "me@email.com"
emailsendto(1) = "example@example.com"
SendEmail "Hello", "Good Morning!", emailsendto</code></pre>


<li><strong>SendEmailString(subject As String, body As String, emails As String)</strong></li>

<pre><code>SendEmailString "Hello", "Good Morning!", "me@email.com,example@example.com"</code></pre>


<li><strong>SendEmailStringCC(subject As String, body As String, emails As String, Optional emailCC As String = "", Optional emailBCC As String = "")</strong></li>

<pre><code>SendEmailStringCC "Hello", "Good Morning!", "me@email.com,example@example.com", "emailcc@copyto.com", "emailbcc@blindcopyto.com"</code></pre>


<li><strong>SendEmailStringCCAttach(subject As String, body As String, emails As String, Optional emailCC As String = "", Optional emailBCC As String = "", Optional attachment As String = "")</strong></li>

<pre><code>SendEmailStringCCAttach "Hello", "Good Morning!", "me@email.com,example@example.com", "emailcc@copyto.com", "emailbcc@blindcopyto.com", "C:\folder1\folder2\file.txt"</code></pre>


<li><strong>SendEmailStringHTML(subject As String, body As String, emails As String, Optional emailscc As String, Optional emailsbcc As String, Optional attachment As String, Optional signature As Boolean = False)</strong></li>

<pre><code>SendEmailStringHTML "Hello", "&lt;html&gt;&lt;body&gt;&lt;font size=""+5"" color=""red""&gt;Good Morning!&lt;/font&gt;&lt;/body&gt;&lt;/html&gt;", _
  "me@email.com,example@example.com", "emailcc@copyto.com", "emailbcc@blindcopyto.com", "C:\folder1\folder2\file.txt", True</code></pre>

Note: <strong>SendEmailStringHTML</strong> function depends on the <strong>EncodeFile</strong> function (last function on the module). This one requires <strong>Microsoft XML, v6.0 (or v3.0)</strong> reference.
