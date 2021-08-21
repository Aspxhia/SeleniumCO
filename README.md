# SeleniumCO
<b>SeleniumCO made for Cloudoffix.</b>

## What is SeleniumCO ?
<b>SeleniumCO,sends your timesheet to the person (manager) you want by e-mail</b>

## How to use SeleniumCO
<table style="width:100%">
  <tr>
    <th>Method</th>
    <th>Using for</th>
  </tr>
  <tr>
    <td>LoginCO()</td>
    <td>Build your driver and login Cloudoffix login website</td>
  </tr>
  <tr>
    <td>AllList()</td>
    <td>All data about on your timesheet</td>
  </tr>
    <tr>
    <td>DateList(AllList())</td>
    <td>Return Turkish day names on your timesheet</td>
  </tr>
  </tr>
  <tr>
    <td>TextList(AllList())</td>
    <td>Return Descriptions on your timesheet</td>
  </tr>
  <tr>
    <td>MixedList(DateList(),TextList())</td>
    <td>Return Turkish day names and Descriptions on your timesheet</td>
  </tr>
  <tr>
    <td>SendMail(To,Subject,Body,CC)</td>
    <td>Sending Mail with your outlook. You dont need any SMTP Server for sending mail with SeleniumCO!</td>
  </tr>
    <tr>
    <td>MixedListSendMail(MixedList(),FirstSentence,To,Subject,CC)</td>
    <td>Sending Mail your Turkish day names and Descriptions with your outlook. You dont need any SMTP Server for sending mail with SeleniumCO!</td>
  </tr>
</table>

## Which Libraries are used
selenium 3.141.0</br>
time</br>
bs4 4.9.3</br>
datetime</br>
pywin32 301</br>


