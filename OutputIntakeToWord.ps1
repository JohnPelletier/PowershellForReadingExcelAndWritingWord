#
#
# THIS POWERSHELL CODE TAKES an EXCEL FILE AND EXTRACTS COLUMN INFO THAT IT THEN FORMATS FOR PRODUCING A WORD DOCUMENT
# In this particular use, I was receiving xlsx files that had lists of requested features and I wanted an automated way to show their status in a monthly document and email.
# However this could easily be adapted to other use cases that aren't requirements.
# The hard part of this was the formatting of the document, which is why I offer it as public source
# The basic format was an intro, a quick jump list to specific products and then the status of each team's candidate requirements, then a footer at the end.
#
# Overall constant settings for this module

#This is the main input file, looking for particular columns of information
$fileName = "IntakeReport.xlsx"

#This is a footer file that will be added to the end of the new docx. It's just a way to append a footer to any word doc that you write
$footerInputFileName = "C:\Intake Email Footer.docx"

#these fields were for troubleshooting back when I was developing this for the first time
$troubleshoot = $false
$limitRows = 0 #set to 0 means don't limit

#this avoids getting tripped up by blank cells in the spreadsheet
$noLink = "NoLink"

#These are different statuses for the candidate requirements in the excel file
$statusOrder = 
    "Already Done",
    "Implementing Immediately (breaks plan)", 
    "Candidate in 1.4",
    "Candidate in 2.1", 
    "Candidate in 2.2",
    "Candidate in 2.3",
    "Researching", 
    "Not Approved"

   
        $jiraString = "https://companyname.atlassian.net/browse/"

$whatIsThisEmail = "[Jump to: What is this email?]"
$tocText = "Interested in the decisions for a specific product?"
$specialString = "xx" #string that is used to make a header string temporarily unique and then find it later

#Font colors
$colorGray = 0x888888
$colorBlack = 0x000000
$statusColor = 0x0080FF #actually orange but I like it

#fake names of product managers
$responsiblePMs = @{
    "Product1"  = "Al Alfonse";
    "Product2"  = "Billy Budd";
    "Product3"  = "Chris Cristo";
    "Product4"  = "David Dillard";
    "Product5"  = "Emma Evangelista";
    "Product6"  = "Frank Florence";
    "Product7"  = "Gib Gamma";
    "Product8"  = "Harry Hilliard";
    "Product9"  = "Ignacio Iverson";
    "Product10" = "Jannette Jacobson";
}


#
#This function formats a single request and outputs it to the file 
#
function Write-Requirement {

    param (
        $document,
        $Selection,
        [pscustomobject]$customObject
    )
    
    
    $document.Styles["Normal"].ParagraphFormat.SpaceAfter = 0;
    
    $Selection.Range.ListFormat.ListIndent()
    $Selection.Range.ListFormat.ApplyBulletDefault()

    #<Name of Requirement> (Source: something)
    #
    $Selection.Style = 'Normal';
    $Selection.Font.Name = 'Calibri';
    $Selection.Font.Bold = $true;
    $Selection.Font.Size = 11;
    $Selection.Font.Color = $colorBlack #black
    $Selection.Range.ListFormat.ApplyBulletDefault() #toggle bullets on
    
    $Selection.TypeText($customObject.RequestTitle)
    $Selection.Font.Color = $colorGray #gray    

    #
    #insert rally ticket with link if it exists
    #[DE27465] format
    #
    if ($customObject.RallyTicket) {
        $Selection.TypeText(" [");
        #either put in a linked ticket string or just the string if there is no link
        if ($customObject.RallyLink -ne $noLink) {
            $document.Hyperlinks.Add(
            $Selection.Range,
            $customObject.RallyLink, 
            $null, $null,
            $customObject.RallyTicket)
        } else {
            $Selection.TypeText("{0}" -f $customObject.RallyTicket);  
        }
        $Selection.TypeText("]");
    }

    #
    #print the source with a link if it exists
    #
    $Selection.TypeText(" (Source: ");
    if ($customObject.SourceLink -ne $noLink) {
        $document.Hyperlinks.Add(
                $Selection.Range,
                $customObject.SourceLink, 
                $null, $null,
                $customObject.Source)
    } else {
        $Selection.TypeText("{0}" -f $customObject.Source);  
    }
    $Selection.TypeText(")");
    $Selection.TypeParagraph()
    
    #toggle bullets off
    $Selection.Range.ListFormat.ApplyBulletDefault()
    $selection.Range.ListFormat.ListIndent()
    $Selection.Font.Color = $colorGray #gray    

    #troubleshooting code
    #Description
    #$Selection.TypeText("Product: {0}" -f $customObject.Product);
    #$Selection.TypeParagraph()
    #$Selection.TypeText("Status: {0}" -f $customObject.Status);
    #$Selection.TypeParagraph()
    #end troubleshooting code

    #Description
    $Selection.Font.Bold = $true;
    $Selection.TypeText("Description: ");

    $Selection.Font.Bold = $false;
    $Selection.TypeText($customObject.Description);
    $Selection.TypeParagraph()

    #Reason
    $Selection.Font.Bold = $true;
    $Selection.TypeText("Reason: ");

    $Selection.Font.Bold = $false;
    $Selection.TypeText($customObject.Reason);

    $Selection.TypeParagraph()
    $Selection.TypeParagraph()
}


#
# STATUS HEADER (Not Approved, Candidate in 21.x, Researching)
#
function Write-Status-Header {

    param (
        $document,
        $Selection,
        [string] $status
    )

    $document.Styles["Normal"].ParagraphFormat.SpaceAfter = 0;
    #$p = $document.Paragraphs.Add();

    #$Selection.ListFormat.ListIndent()
    
    #
    #Candidate in 21.x format
    #
    $Selection.Style = 'Normal';
    $Selection.Font.Name = 'Calibri';
    $Selection.Font.Bold = $true;
    $Selection.Font.Size = 15;

    $Selection.Font.Color = $statusColor #


    $Selection.TypeText("{0}" -f $customObject.Status);  
    $Selection.TypeParagraph()

    if ($troubleshoot) {
        write-output ""
        write-output ("STATUS: {0}" -f $status)
    }

}

#
#Write product header (Designer, Server, Engine, etc.)
#
function Write-Product-Header {

    param (
        $document,
        $Selection,
        [string] $product,
        [pscustomobject] $customObject
    )

    #This adds a bookmark so it's easy to find this product from the table of contents
    #bookmarks won't work with spaces in the name so remove spaces
    $product = $customObject.Product.Replace(' ','')

    $Selection.Bookmarks.Add($product)
    $Selection.Style = 'Normal';
    $Selection.Font.Name = 'Calibri';
    $Selection.Font.Bold = $true;
    $Selection.Font.Size = 26;

    $Selection.Font.Color = 0x000000 #black

    #$Selection.Range.ListFormat.ListOutdent()
    
    $Selection.Range.ListFormat.ApplyNumberDefault()
    $Selection.Range.ListFormat.ListOutdent()
    $Selection.TypeText("{0} " -f $customObject.Product);

    $Selection.Font.Bold = $true;
    $Selection.Font.Size = 11;

    $Selection.Font.Color = 0x888888 #gray

    $Selection.TypeText("   PM Lead(s): {0}" -f $responsiblePMs[$customObject.Product]);  
    $Selection.TypeParagraph()

    #if ($troubleshoot) {
    #    write-output ""
    #    write-output ("PRODUCT: {0}" -f $product) # Responsible PM: {1}" -f $product)
    #}

}


#
#Write-Doc-Header-Text just adds the raw text to start the document
#  (Note: There is another function to go back and add bookmarks)
#
function Write-Doc-Header-Text {
    param (
        $document,
        $Selection,
        $whatIsThisEmail,
        [string[]] $bookmarks
    )

    #calibri 11 with link to bookmark "Explanation"
    $Selection.Style = 'Normal';
    $Selection.Font.Name = 'Calibri';
    $Selection.Font.Bold = $true;
    $Selection.Font.Size = 11;
    #$Selection.Font.Underline = $true;
     
    #add the text to the selection
    $Selection.TypeParagraph()

    $Selection.TypeText($whatIsThisEmail);
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()

    $Selection.Font.Bold = $true;
    $Selection.Font.Size = 14;
    $Selection.TypeText($tocText);
    $Selection.TypeParagraph()

    $index = 1

    #Indent the table of contents so it falls under the tocText
    #$Selection.Range.ListFormat.ListIndent()
    $Selection.Font.Size = 11;
    $bookmarks | ForEach-Object {
        $bm = $_+$specialString #the specialstring marks it as a future link text so it's easier to find after processing requests
        $Selection.TypeText("  {0}. " -f $index) 
        $Selection.TypeText(" {0}" -f  $bm)
        $Selection.TypeParagraph()
        $index +=1 
    }

    $Selection.TypeParagraph()
    $Selection.TypeParagraph()
    #$Selection.Range.ListFormat.ListOutdent()

    #this should toggle numbering off

}

#
#Finds the indicated string in the document and returns a range to it, or null if it doesn't find it
#
function SearchForText {
    param (
        [Microsoft.Office.Interop.Word.DocumentClass]    $document,
        [Microsoft.Office.Interop.Word.ApplicationClass] $word,
        [string]$searchText
    )

    #Include the entire text in the search
    $document.Content.Select()

    #search for the first string
    $word.Selection.Find.Text = $searchText
    if ($word.Selection.Find.Execute()) {
        Write-Debug ("{0} is at {1} and ends at {2}" -f $searchText, $word.Selection.Range.Start, $word.Selection.Range.End)
        $range = $word.Selection.Range
    } else {
        $range = $null

    }
    return $range
}



###############################################################################################################
#   START OF SCRIPT CODE
#   A LITTLE MESSY BUT IT WORKED LIKE A CHARM FOR 2 YEARS!
##############################################################################################################

#Open a new word document
$word = New-Object -ComObject "Word.Application"
$word.Visible = $true
$document = $word.Documents.Add();
$Selection = $word.Selection


# Get the excel input open
$xl = New-Object -ComObject Excel.Application 
$workbook = $xl.Workbooks.open($fileName)

#assumes no blank rows, so clean up first
$rowCount = $workbook.ActiveSheet.UsedRange.Rows.Count

#for troubleshooting purposes to speed things up, limit rows of requirements that you process
if ($limitRows -gt 0) {
    $rowCount = $limitRows
}


$columnNames = $workbook.ActiveSheet.Rows | select -first 1 | where Value2 -ne $null
$columnCount = ($columnNames.Value2 | Select).Count


#Neat trick to have a cross reference to smartly figure out the columns correctly on each run
$lookupRec = @{}
$i = 1
while ($i -le $columnCount) {
    $lookupRec[$columnNames.Value2[1,$i]] = $i  
    $i++;
}

# Get the active sheet of the excel open
$rows = $workbook.ActiveSheet.Rows | select -first ($rowCount-1) -skip 1 

#This does a primary and secondary sort, Primary: Team, Secondary: Status of candidate requests
#This is my chance to sort in a different order for Product and then custom order for status
#first do a custom sort by the secondary key
$sortedRows = $rows | Sort-Object { 
    $_.Columns[$lookupRec["Primary"]].Value2,   
    $statusOrder.IndexOf($_.Columns[$lookupRec["Status"]].Value2) 
}


#keep track of which status (Candidate for xx.xx) and which product (Designer, Server, etc) you're on so you know when you've switched
$status = ""
$product = ""

#set up a list of bookmarks that we need to make the table of contents at the top, listing the teams/products
$bookmarks = @()
$rows | ForEach-Object {
    $bookmarks += $_.Columns[1].Value2
}
$bookmarks = $bookmarks | select -Unique | Sort-Object

#Add an explanation header and a link to a bookmark
Write-Doc-Header-Text $document $Selection $whatIsThisEmail $bookmarks

#loop through all rows to format their requests
$sortedRows | ForEach-Object {
    
    $customObject = [pscustomobject]@{

        Product      = $_.Columns[$lookupRec["Primary"]].Value2
        Status       = $_.Columns[$lookupRec["Status"]].Value2
        #ProductMgr   = $_.Columns[$lookupRec["Responsible PM"]].Value2  <note at some point, we decided to do this different since some teams had two PMs>
        RequestTitle = $_.Columns[$lookupRec["Request Idea"]].Value2
        Description  = $_.Columns[$lookupRec["Description"]].Value2
        RallyTicket  = $_.Columns[$lookupRec["Issue Ticket (optional)"]].Value2
        Reason       = $_.Columns[$lookupRec["Reason"]].Value2
        Source       = $_.Columns[$lookupRec["Source"]].Value2
        SourceLink   = $_.Columns[$lookupRec["Source Link (optional)"]].Value2
        RallyLink    = $noLink
        RallySubAddr = $noLink
    }

    #clean up a null source link
    if ($customObject.SourceLink -eq $null) { $customObject.SourceLink = $noLink }


    #Rally links come from excel in a strange format and need special formatting
    $i = $lookupRec["Issue Ticket (optional)"]

    #try to query the JIRA ticket value to add the hyperlink based on the ticket number
    if ($_.Columns[$i].Text) {
        if ($_.Columns[$i].Hyperlinks.Count) {
            #Non-Rally: this code will probably work for issue links other than Rally (Git, etc)
            $trimString = $_.Columns[$i].Hyperlinks[1].SubAddress
            $customObject.RallyLink = $_.Columns[$i].Hyperlinks[1].Address + "#" + $trimString 
        }
    }

    #Source links are assumed to be simply a combination of Address and SubAddress
    #if the explicit field for SourceLink is empty but the source itself has a hyperlink...
    $i = $lookupRec["Source Link (optional)"]
    if ($_.Columns[$i].Hyperlinks.Count -and $customObject.SourceLink -eq $null) {
        #grab the link from the source
        $customObject.SourceLink = $_.Columns[$i].Hyperlinks[1].Address.TrimEnd('/') + $_.Columns[$i].Hyperlinks[1].SubAddress
    }

    #Detect if this next request comes for a different product in the sorted list
    if ($product -ne $customObject.Product) {
        $product = $customObject.Product
        $status = "" #need to reset this for each new product

        #add code to print the new product
        Write-Product-Header $document $Selection $product $customObject
    }

    #this code assumes the rows are grouped by product and then status
    if ($status -ne $customObject.Status) {
        $status = $customObject.Status

        #add code to print the status
        Write-Status-Header $document $Selection $status
    }

    #output
    Write-Requirement $document $Selection $customObject
    
}


#insert footer information at the end of the file
$word.Selection.InsertFile($footerInputFileName)


#Create hyperlinks for bookmarks to each product
$bookmarks | Foreach {
        $sstring = $_+$specialstring
        $linkRange = SearchForText $document $word ($sstring)
        
        $bmarkString = $_.Replace(' ', '')
        $bmark = $document.Bookmarks.Item($bmarkString)

        $document.Hyperlinks.Add(
        $linkRange,
        $null, 
        $bmark, #bookmarks go in the subaddress
        $null,
        $_)
}

#One last bookmark link to make...
$bmrange = SearchForText $document $word 'FAQ'
$linkrange = SearchForText $document $word $whatIsThisEmail

#Add the link from the top to the FAQ section at the bottom
$document.Hyperlinks.Add(
        $linkrange,
        $null, 
        $bmrange, #bookmarks go in the subaddress
        $null)


$OutputFileName = "outputfile.docx"
$document.SaveAs($OutputFileName)

#have to close things out or it gets messy
$workbook.Close()
$xl.Quit()
$word.Quit()
