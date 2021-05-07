# New-TemplateChooser
New Template Chooser
Template-Chooser

Templates Choose is to chooser word templates is storeted in Sharepoint of the company, the document display to the user choose. In Office 365 projects, SharePoint or document structure planning, the question often araises, where to store templates so that every user has access to them and, more importantly, always uses the latest version. Therefore, here is a tool powered approach that solves this problem and uses cloud technology to solve all these problems.

## The second experience with Add-in

This Add-In and I made using Visual Studio and in the same language Javascript This case selected the document and display it in the Word document, To make a chooser it will be subtitle and they click the button and each of them will display different templates there is a All button to display all of them in a single selection and easier the user will select them, depends what type of document they are looking to display in Microsoft Word.

I code this function to display different subjects of templates in a list, Also I code the view of the Add-in using the tools of CSS and HTML for the looking of the tempales.

Part of the CSS Code I use for making the display of the templates with a button and making the background

Working in the Back end to connected to the Sharepoint is coming the code and the Template update it.

Share a folder In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog. Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose Properties. Within the Properties dialog window, open the Sharing tab and then choose the Share button.

## Deploy it in Azure
Making a web application on Azure for the API and using back end code I will get the successful connection, using conceptes as MVC for my model and making my life easier and understandoble my code.

## html

the body of the HTML has <script> where I indicate the src of the file of the js that will containe the html.

<script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script> <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
<script src="FunctionFile.js" type="text/javascript"></script>
Javascript the next fuction is an example of the display template:

function displaytemplates() {
    var templates = ['Templatechooser.docx', 'Template2.docx'];
    templates = new docxTemplater();
   templates.loadZip(zip);
    //forlook for the image
    for (var i = 0; i < templates.length; i++) {
        var File = templates[i];
        //add-in container for display the imagine with the url and the class html addin 
        $(".templates").append(
            '<div class= "tn">' +
            '<a" http://TemplateChooserWeb/Templates' + File + '" alt = "templates" /> ' +
            '</div>'
        );
    }
}
