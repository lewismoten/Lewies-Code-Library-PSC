// Setup the following attributes
var Columnizer_VisibleColumns = 4;
var Columnizer_ColumnHeight = 400;
var Columnizer_Padding = 8;
var Columnizer_FontSize = 10; // Make sure ColumnHeight can be evenly divided by font-size 
var Aritlce_FontFamily = "Arial";

// These attributes are for internal use.  Do not modify.
var Columnizer_Html = "";
var Columnizer_FullHeight = 0;
//var Columnizer_Columns = 0;

function Columnizer_ShowPage(page)
{

	// positions text within each column so that correct portion is
	// displayed according to the page number

	var StartColumnIndex = ((page-1) * Columnizer_VisibleColumns);

	for(var i = 0; i < Columnizer_VisibleColumns; i++)
	{
		var ColumnText = document.getElementById("Columnizer_ColumnText" + (i+1))
		if(ColumnText != null)
		{		
			ColumnText.style.top = -1 * (StartColumnIndex + i) * Columnizer_ColumnHeight;
		}
	}
}
function Columnizer_PageChanged()
{
	// handles the page listbox onchange event
	Columnizer_ShowPage(this.selectedIndex + 1);
}
function Columnizer_PrepareHtml(html)
{
	html += "&nbsp;";
	html = html.replace(/<H\d[^>]*>/ig,"<DIV style='font-weight:bold'>");
	html = html.replace(/<\/H\d>/ig,"</DIV>");
	return html;
}
function Columnizer_Initialize(theArticle)
{
	//document.getElementById("Article")

	// sets up the article for the first time.
	// takes all existing html out of the article object and caches it.
	// then, creates children for each column and positions the text within
	// them.  Also builds a select box to navigate each page.

	Columnizer_Html = "";
	Columnizer_FullHeight = 0;
	Columnizer_Columns = 0;
	var FullWidth = theArticle.offsetWidth-Columnizer_VisibleColumns;
	
	Columnizer_Html = Columnizer_PrepareHtml(Article.innerHTML);
	theArticle.innerHTML = "";
	
	for(var ColumnIndex = 0; ColumnIndex < Columnizer_VisibleColumns; ColumnIndex++)
	{
		var Column = document.createElement("SPAN");
		Column.setAttribute("id", "Columnizer_Column" + (ColumnIndex + 1))
		Column.style.width =  Math.floor(FullWidth / Columnizer_VisibleColumns) + "px";
		Column.style.height = Columnizer_ColumnHeight + "px";
		Column.style.overflow = "hidden";
		if(ColumnIndex != 0)
		{
			Column.style.borderLeft = "1px solid black";
		}
		theArticle.appendChild(Column);

		var Text = document.createElement("SPAN");
		Text.innerHTML = Columnizer_Html;
		Text.setAttribute("id", "Columnizer_ColumnText" + (ColumnIndex + 1));
		
		// Scroll column ...
		//Text.style.top = -1 * Index * Columnizer_ColumnHeight;
		
		Text.style.display = "block";
		Text.style.textAlign = "justify";
		Text.style.position = "absolute";
		Text.style.fontFamily = Aritlce_FontFamily;
		Text.style.paddingLeft = Columnizer_Padding + "px";
		Text.style.paddingRight = Columnizer_Padding + "px";
		
		 
		// make sure size and lineHeight divide
		// evenly with Columnizer_ColumnHeight
		Text.style.fontSize = Columnizer_FontSize + "px";
		Text.style.lineHeight = Columnizer_FontSize + "px";
		 
		Column.appendChild(Text);

		if(Columnizer_FullHeight == 0)
		{
			Columnizer_FullHeight = Text.offsetHeight;
		}
	}

	var ColumnCount = Math.floor(Columnizer_FullHeight / Columnizer_ColumnHeight);
	if(Columnizer_FullHeight % Columnizer_ColumnHeight != 0)
	{
		ColumnCount ++;
	}

	theArticle.appendChild(document.createElement("BR"))

	var PageList = document.createElement("SELECT");
	PageList.setAttribute("id", "Columnizer_lstPages");
	theArticle.appendChild(PageList);
	var Pages = Math.floor(ColumnCount / Columnizer_VisibleColumns);
	if(ColumnCount % Columnizer_VisibleColumns != 0)
	{
		Pages++;
	}
	
	for(var Page = 1; Page <= Pages; Page++)
	{
		var Option = document.createElement("OPTION");
		Option.innerHTML = "Page " + Page;
		PageList.appendChild(Option);
	}
	
	PageList.onchange = Columnizer_PageChanged;

	Columnizer_ShowPage(1);
} 