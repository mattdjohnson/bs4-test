from docx import Document
from htmldocx2 import *

document = Document()
new_parser = HtmlToDocx()
# do stuff to document

html_doc = """<div id="document_content"><p><br></p>
<div class="Section1">
<h1>Hi I am an H1</h1>
<h2>Hi I am an H2</h2>

    <ol style="list-style-position: inside;">
        <li>&nbsp;FIRST ITEM<br /></li>
        <li>&nbsp;SECOND ITEM</li>
    </ol>

    <ul style="list-style-position: inside;">
        <li>ONE</li>
        <li>TWO</li>
    </ul>
<hr>
<p style="text-align: center;" ><span style="background-color:rgb(186,218,85);" >TITLE</span></p>
<p class="MsoTitle" style="text-align: center;color:red"><strong>SYSTEM SECURITY PLAN (SSP)</strong></p>
<p class="MsoSubtitle"><strong>National Reconnaissance Office</strong></p>

<p class="MsoNormal"><span style="font-size: 12pt;">Project ID: <strong>200410</strong></span></p>

<p class="MsoNormal"><span style="font-size: 12pt;">Project Name: <strong id="project_id">Unclassified Enterprise Continuous Monitoring</strong></span></p>

<p class="MsoNormal"><span style="font-size: 12pt;">Date: <strong><span>08-23-2021</span></strong></span></p>
</div>
<p>< <br style="mso-special-character: page-break; page-break-before: always;" clear="all"></p>
<div class="Section2">
<p class="MsoNormal"><span style="font-size: 12pt;"><strong><em>Introduction</em></strong></span></p>

<p class="MsoNormal"><span style="font-size: 12pt;"><em>(U//FOUO) The completion of system security plans is a requirement of the Office of Management and Budget (OMB) Circular A-130, Management of Federal Information Resources and Public Law 107-347, the Federal Information Security Management Act (FISMA), NIST Special Publication 800-37 Rev.&nbsp; 2, Guide for Applying the Risk Management Framework to Federal Information Systems, and Committee on National Security Systems (CNSS) Instruction (CNSSI) 1253, Security Categorization and Control Selection for National Security Systems. Federal agencies are required to identify each computer system that contains sensitive information, and to prepare and implement a plan for the security and privacy of these systems.&nbsp; The objective of system security planning is to improve protection of information technology (IT) resources.&nbsp; The protection of a system is documented in a System Security Plan (SSP).</em></span></p>
<br
<p class="MsoNormal"><span style="font-size: 12pt;"><em>(U//FOUO) The SSP documents the results of planning and implementing adequate, cost-effective security protection for a system.&nbsp; It reflects input from management responsible for the system, including the information system owner, information owners, the system operator, the system security manager, and system administrators.&nbsp; The SSP delineates responsibilities and expected behavior of all individuals who access the system. </em></span></p>
<br>
<p class="MsoNormal"><span style="font-size: 12pt;"><em>(U//FOUO) The purpose of the SSP is to provide an overview of the security of the information system and to describe the controls and critical elements in place or planned for, based on CNSS Instruction 1253.&nbsp; The SSP provides sufficient information to enable an understanding of the implementation of each security control in the context of the information system.&nbsp; The System Security Plan also contains as supporting appendices or as references to appropriate sources, other key security-related documents such as a risk assessment, privacy impact assessment, system interconnection agreements, contingency plan, security configurations, configuration management plan, and incident response plan.&nbsp; This SSP template follows guidance contained in NIST Special Publication 800-37 Revision 2. </em></span></p>
</div>
<p>< <br style="mso-special-character: page-break; page-break-before: always;" clear="all"></p>

<p class="MsoTocHeading"><strong>TABLE OF CONTENTS</strong></p>

<table style="border: 1px solid gray;font-family:'Times New Roman';max-width:5in !important;" id="x_snc_authorizatio_stakeholders_child_iso_name" class="table table-striped table-bordered table-condensed"><thead><tr><th style="border: 1px solid gray;"> User</th><th style="border: 1px solid gray;"> Role</th><th style="border: 1px solid gray;"> Employee Type</th><th style="border: 1px solid gray;"> Email</th></tr></thead><tbody><tr style="border: 1px solid gray;"><td style="border: 1px solid gray;"> ISO Test</td><td style="border: 1px solid gray;"> ISO</td><td style="border: 1px solid gray;"></td><td style="border: 1px solid gray;"> Joshua.hill@recro.com</td></tr><tr style="border: 1px solid gray;"><td style="border: 1px solid gray;"> Joshua Hill</td><td style="border: 1px solid gray;"> ISO</td><td style="border: 1px solid gray;"></td><td style="border: 1px solid gray;"> joshua.hill@recro.com</td></tr><tr style="border: 1px solid gray;"><td style="border: 1px solid gray;"> Taylor Wirth</td><td style="border: 1px solid gray;"> ISO</td><td style="border: 1px solid gray;"> GOV</td><td style="border: 1px solid gray;"> taylor.wirth@recro.com</td></tr></tbody></table>

</div>
</div>
"""
new_parser.add_html_to_document(html_doc, document)

# do more stuff to document
document.save('htmldocx.docx')