$remaining = 0
$username = "John Doe"

   if ($remaining -lt 15) {
   
       if ($remaining -lt 1) {
           $display = "Your password has expired."
           } else {
           $display = "Your password will expire in $remaining days."
           }
       $MyVariable = @"
Dear $UserName<br/><br/>

$display  If you fail to change it you will not be able to connect to College resources.<br/>
This includes Gateway, ANGEL, Clean Access, and the wireless network, to name a few.<br/><br/>

If you are on campus, follow the steps below to change your password:<br/><br/>

- Log into a College-owned machine, such as a lab machine<br/>
- Press CTRL+ALT+DEL<br/>
- Click Change a Password...<br/>
- Remember that passwords have to be at least 8 characters long, and cannot be any of your previous 5 College passwords<br/><br/>

If you are off campus, follow the steps below to change your password:<br/><br/>

- Please click <a href="https://gateway.manchester.edu">here</a> to log into Gateway and reset your password.  The reset link is at the top of the page<br/>
- Remember that passwords have to be at least 8 characters long, and cannot be any of your previous 5 College passwords<br/><br/>

If your password has already expired:<br/><br/>

- Reply to this email and ask for a password reset<br/>
- Or come by the Help Desk between 8 and 5 M-F with a photo ID<br/><br/><br/><br/>
 


Thank you,<br/><br/>
Manchester College ITS Help Desk<br/>
260-982-5454<br/>
<a href="mailto:helpdesk@manchester.edu">helpdesk@manchester.edu</a><br/>
<a href="https://helpdesk.manchester.edu"></a>https://helpdesk.manchester.edu<br/><br/>

ITS will <strong>never</strong> ask you for your password.
"@

    send-mailmessage -to helpdesk@manchester.edu -from helpdesk@manchester.edu -Subject "Manchester ITS:  Password expiration notice" -body $MyVariable  -smtpserver 10.90.254.73 -BodyAsHtml
}