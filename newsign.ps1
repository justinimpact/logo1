<#
Created:	 2015-02-22
Version:	 1.0
Author       Peter Löfgren
Homepage:    http://syscenramblings.wordpress.com

Disclaimer:
This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the authors or DeploymentArtist.

Author - Peter Löfgren
    Twitter: @LofgrenPeter
    Blog   : http://syscenramblings.wordpress.com
#>

#Find the User and all values

#Copy-Item -Source https://raw.githubusercontent.com/justinimpact/logo1/master/image001.png
#-Destination C:\logo3.png
# https://impactittech.sharepoint.com/:i:/s/TheDreamTeam/EYcZLa3GncdLl300P0-ooHkB4WydAEIah7UrWHzq-ie2hA?e=uhzU2g

Invoke-WebRequest https://raw.githubusercontent.com/justinimpact/logo1/master/image001.png -OutFile C:\logo4.png

$UserName = $env:username
$Logo = "c:\logo4.png"
$SignatureName = "winslow"


$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()
$ADDisplayName = $ADUser.DisplayName
$ADTitle = 'Mr'
$ADFirstname = 'Cathya'
$ADSecondname = 'Djanogly'
$ADjobtitle = 'Tax Lawyer'
$ADemail = 'cathya@winslows.co.uk'
$ADmobilephone = '+44 (0) 750 090 4796'
$ADtelbelfast = '+44 (0) 289 521 6744'
$ADtellondon = '+44 (0) 203 196 5582'
$ADMobile = $ADUser.Mobile
$ADFax = $ADUser.facsimileTelephoneNumber
$AdPhone = $ADUser.telephoneNumber
$AdCompany = 'Dude 123'
$AdStreet = $ADUser.streetAddress
$AdZip = $ADUser.postalCode
$AdLocation = $ADUser.l
$AdCountry = $ADUser.co
$ADemail = $ADUser.mail
$ADBox = $ADUser.postOfficeBox


#Create the signaturefile
$Html=@"
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body lang=EN-GB link=blue vlink="#954F72" style='tab-interval:36.0pt'>

<div class=WordSection1>

<table class=MsoTableGrid border=0 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;border:none;mso-yfti-tbllook:1184;mso-padding-alt:
 0cm 5.4pt 0cm 5.4pt;mso-border-insideh:none;mso-border-insidev:none'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
  <td width=217 valign=top style='width:163.05pt;padding:0cm 5.4pt 0cm 5.4pt'>
  <p class=MsoNormal><a name="OLE_LINK10"></a><a name="OLE_LINK5"></a><a
  name="OLE_LINK4"></a><a name="OLE_LINK8"></a><a name="OLE_LINK7"></a><a
  name="_MailAutoSig"><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:minor-fareast;
  mso-fareast-language:EN-GB;mso-no-proof:yes'>$ADFirstname $ADSecondname<o:p></o:p></span></span></span></span></span></span></a></p>
  <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><b style='mso-bidi-font-weight:normal'><span
  style='mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:minor-fareast;
  mso-fareast-language:EN-GB;mso-no-proof:yes'>$ADjobtitle<o:p></o:p></span></b></span></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='font-size:10.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>Tel. (Belfast):&nbsp;$ADtelbelfast<o:p></o:p></span></span></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='font-size:10.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>Tel. (London):&nbsp;$ADtellondon<o:p></o:p></span></span></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='font-size:10.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>Mobile: $ADmobilephone<o:p></o:p></span></span></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='font-size:10.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>Email: $ADemail<o:p></o:p></span></span></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='font-size:10.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>Website: </span></span></span></span></span></span></span><a
  href="http://www.winslows.co.uk/"><span style='mso-bookmark:_MailAutoSig'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='font-size:10.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>www.winslows.co.uk</span></span></span></span></span></span></span></a><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='mso-fareast-font-family:"Times New Roman";
  mso-fareast-theme-font:minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:
  yes'><o:p></o:p></span></span></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span></span></span>
  <td width=384 valign=top style='width:287.75pt;padding:0cm 5.4pt 0cm 5.4pt'>
  <p class=MsoNormal><!--[if gte vml 1]><v:shapetype id="_x0000_t75"
   coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe"
   filled="f" stroked="f">
   <v:stroke joinstyle="miter"/>
   <v:formulas>
    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
    <v:f eqn="sum @0 1 0"/>
    <v:f eqn="sum 0 0 @1"/>
    <v:f eqn="prod @2 1 2"/>
    <v:f eqn="prod @3 21600 pixelWidth"/>
    <v:f eqn="prod @3 21600 pixelHeight"/>
    <v:f eqn="sum @0 0 1"/>
    <v:f eqn="prod @6 1 2"/>
    <v:f eqn="prod @7 21600 pixelWidth"/>
    <v:f eqn="sum @8 21600 0"/>
    <v:f eqn="prod @7 21600 pixelHeight"/>
    <v:f eqn="sum @10 21600 0"/>
   </v:formulas>
   <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
   <o:lock v:ext="edit" aspectratio="t"/>
  </v:shapetype><v:shape id="Picture_x0020_1" o:spid="_x0000_s1026" type="#_x0000_t75"
   alt="Winslows_EmailLogo" style='position:absolute;margin-left:55.8pt;
   margin-top:3.6pt;width:157.5pt;height:112.5pt;z-index:251659264;
   visibility:visible;mso-wrap-style:square;mso-wrap-distance-left:9pt;
   mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;
   mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;
   mso-position-horizontal-relative:text;mso-position-vertical:absolute;
   mso-position-vertical-relative:page'>
   <v:imagedata src="$Logo" o:title="Winslows_EmailLogo"/>
   <w:wrap anchory="page"/>
  </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;
  position:absolute;z-index:251659264;margin-left:74px;margin-top:5px;
  width:210px;height:150px'><img width=210 height=150
  src="$Logo" alt="Winslows_EmailLogo" v:shapes="Picture_x0020_1"></span><![endif]><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span></span></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:minor-fareast;
  mso-fareast-language:EN-GB;mso-no-proof:yes'><span
  style='mso-spacerun:yes'>                         </span></span></span></span></span></span></span><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='font-size:10.0pt'><o:p></o:p></span></span></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span></span></span>
 </tr>
 <tr style='mso-yfti-irow:1'>
  <td width=217 valign=top style='width:163.05pt;padding:0cm 5.4pt 0cm 5.4pt'>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='font-size:10.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'>40 Grosvenor
  Gardens, Belgravia, <o:p></o:p></span></span></span></span></span></span></p>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='font-size:10.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'>London SW1W 0EB<o:p></o:p></span></span></span></span></span></span></p>
  <span style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span></span></span>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><b
  style='mso-bidi-font-weight:normal'><span style='mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></b></span></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span></span></span>
  <td width=384 valign=top style='width:287.75pt;padding:0cm 5.4pt 0cm 5.4pt'><span
  style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span></span></span>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK7'><span
  style='mso-bookmark:OLE_LINK8'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:minor-fareast;
  mso-fareast-language:EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK7'><span style='mso-bookmark:OLE_LINK8'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span></span></span>
 </tr>
 <tr style='mso-yfti-irow:2'>
  <td width=601 colspan=2 valign=top style='width:450.8pt;padding:0cm 5.4pt 0cm 5.4pt'>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><a
  name="OLE_LINK9"><span style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";
  mso-fareast-theme-font:minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:
  yes'>Winslows Tax Law Limited is a specialist tax practice with its
  registered office at&nbsp;12-18 Grosvenor Gardens, London, England, SW1W 0DH.
  &nbsp;Winslows Tax Law Limited is a company limited by shares, incorporated
  and registered in England and Wales under number:&nbsp;10566246.&nbsp; We are
  authorised and regulated by the Solicitors Regulation Authority (SRA No:
  637211, link: www.sra.org.uk).<o:p></o:p></span></a></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><span style='font-size:8.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><span style='font-size:8.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>A list of directors and their professional qualifications
  is open to inspection at our registered offices.&nbsp; The word “partner” is
  used to refer to a director of the company, or an employee or consultant who
  is a lawyer with equivalent standing and qualifications.<o:p></o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><span style='font-size:8.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><span style='font-size:8.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>The contents of this e-mail (including any
  attachments) are confidential and may be legally privileged. If you are not
  the intended recipient of this e-mail, any disclosure, copying, distribution
  or use of its contents is strictly prohibited, and you should please notify
  the sender immediately and then delete it (including any attachments) from
  your system.<o:p></o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><span style='font-size:8.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'>Notice</span></b></span></span></span></span><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='mso-bookmark:OLE_LINK9'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'>: the firm does
  not accept service by e-mail of court proceedings, other processes or formal
  notices of any kind without specific prior written agreement.<o:p></o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></b></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'>Notice</span></b></span></span></span></span><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'><span style='mso-bookmark:OLE_LINK9'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'>: As a result of the
  increased risk posed by cyber fraud and especially those relating to bank
  account details, please note that Winslows Tax Law Ltd bank account details
  will NOT change during the course of a matter. Please be vigilant and ensure
  caution is exercised when opening any emails, attachments or links and when
  responding to any requests for your bank account details. We will not accept
  responsibility if you transfer money into an incorrect bank account. <o:p></o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><span style='font-size:8.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></span></p>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='mso-bookmark:OLE_LINK9'><span style='font-size:8.0pt;mso-fareast-font-family:
  "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
  EN-GB;mso-no-proof:yes'>The firm takes reasonable steps to ensure that
  e-mails and attachments do not contain viruses but recipients are responsible
  for ensuring that their systems are protected by appropriate firewalls and
  anti-virus software.<o:p></o:p></span></span></span></span></span></p>
  <span style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
  <p class=MsoNormal><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
 </tr>
 <tr style='mso-yfti-irow:3'>
  <td width=601 colspan=2 valign=top style='width:450.8pt;padding:0cm 5.4pt 0cm 5.4pt'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
 </tr>
 <tr style='mso-yfti-irow:4'>
  <td width=601 colspan=2 valign=top style='width:450.8pt;padding:0cm 5.4pt 0cm 5.4pt'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
 </tr>
 <tr style='mso-yfti-irow:5;mso-yfti-lastrow:yes'>
  <td width=601 colspan=2 valign=top style='width:450.8pt;padding:0cm 5.4pt 0cm 5.4pt'><span
  style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
  <p class=MsoNormal style='text-align:justify'><span style='mso-bookmark:OLE_LINK4'><span
  style='mso-bookmark:OLE_LINK5'><span style='mso-bookmark:OLE_LINK10'><span
  style='font-size:8.0pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
  minor-fareast;mso-fareast-language:EN-GB;mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></span></span></p>
  </td>
  <span style='mso-bookmark:OLE_LINK4'><span style='mso-bookmark:OLE_LINK5'><span
  style='mso-bookmark:OLE_LINK10'></span></span></span>
 </tr>
</table>

<span style='mso-bookmark:OLE_LINK10'></span><span style='mso-bookmark:OLE_LINK5'></span><span
style='mso-bookmark:OLE_LINK4'></span>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

</div>

</body>
</html>
"@


#Create the actuall file
if (!(Test-Path -Path $env:APPDATA\microsoft\signatures)){ mkdir $env:APPDATA\microsoft\signatures }
Copy-Item -Path "$PSScriptRoot\$Logo" -Destination $env:APPDATA\microsoft\signatures -Force

$Html | Out-File "$env:APPDATA\microsoft\signatures\$SignatureName.htm" -Force

#Enforce embedded pictures in outlook
if (!(Test-Path -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail)) { New-Item -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail -ItemType Directory -Force }
New-ItemProperty HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail -Name 'Send Pictures With Document' -Value 1 -PropertyType 4 -Force
if (!(Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail)) { New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail -ItemType Directory -Force }
New-ItemProperty HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail -Name 'Send Pictures With Document' -Value 1 -PropertyType 4 -Force

#Set the signature as default for new mails
New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force
New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force

#Set the signature as default for reply mails
New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force

$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "random.dude@email.com"
$Mail.Subject = "data for Subject"
$Mail.Body ="Example of body..."
$Mail.Send()
