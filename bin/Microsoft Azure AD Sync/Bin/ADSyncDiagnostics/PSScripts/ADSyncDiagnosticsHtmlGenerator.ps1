#-------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation.  All rights reserved.
#-------------------------------------------------------------------------

Function WriteHtml
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $body
    )

    $htmlDoc = [string]::Empty

    #
    # CSS
    #
    $htmlStyle = WriteHtmlStyle

    $htmlHead = WriteHtmlElement -elementName "head" -innerHtml $htmlStyle

    $attributePopupWindow = WriteAttributePopupWindow
    
    $htmlDoc += $attributePopupWindow
    $htmlDoc += "`n"
    $htmlDoc += $htmlHead
    $htmlDoc += "`n"
    $htmlDoc += $body

    $html = WriteHtmlElement -elementName "html" -innerHtml $htmlDoc

    Write-Output $html
}

Function WriteHtmlBody
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $accordion
    )

    $banner = WriteHtmlMessage -message "Microsoft Azure" -color "#00ABEC" -fontSize "36px" -numberOfLineBreaks 0

    $bannerDiv = WriteHtmlElement -elementName "div" -class "banner" -innerHtml $banner

    $heading = WriteHtmlMessage -message $global:HtmlTitle -color "#FFFFFF" -fontSize "30px" -numberOfLineBreaks 0

    $headingDiv = WriteHtmlElement -elementName "div" -class "header" -innerHtml $heading

    $jqueryScript = WriteHtmlElement -elementName "script" -src "https://code.jquery.com/jquery-2.2.4.min.js"

    $jsScript = WriteHtmlScript

    $bodyInnerHtml = [string]::Empty

    $bodyInnerHtml += $bannerDiv
    $bodyInnerHtml += "`n"
    $bodyInnerHtml += $headingDiv
    $bodyInnerHtml += "`n"
    $bodyInnerHtml += $accordion
    $bodyInnerHtml += "`n"
    $bodyInnerHtml += $jqueryScript
    $bodyInnerHtml += "`n"
    $bodyInnerHtml += $jsScript

    $body = WriteHtmlElement -elementName "body" -onload "addRowHandlers()" -innerHtml $bodyInnerHtml

    Write-Output $body
}

Function WriteHtmlAccordion
{
    param
    (
        [string[]]
        [parameter(mandatory=$true)]
        $accordionGroupList,

        [string]
        [parameter(mandatory=$true)]
        $ObjectDn
    )

    $objectDnInnerHtml = [string]::Empty

    $objectDnHeading = WriteHtmlElement -elementName "h2" -innerHtml $global:HtmlObjectDistinguishedNameSectionTitle
    $objectDnValue = WriteHtmlElement -elementName "p" -id "ObjectDnId" -innerHtml $ObjectDn

    $objectDnInnerHtml += $objectDnHeading
    $objectDnInnerHtml += "`n"
    $objectDnInnerHtml += $objectDnValue

    $objectDnDiv = WriteHtmlElement -elementName "div" -innerHtml $objectDnInnerHtml

    $accordionInnerHtml += $objectDnDiv
    $accordionInnerHtml += "`n"
    $accordionInnerHtml += "</br>"
    $accordionInnerHtml += "`n"

    foreach ($group in $accordionGroupList)
    {
        $accordionInnerHtml += $group
        $accordionInnerHtml += "`n"
        $accordionInnerHtml += "</br>"
        $accordionInnerHtml += "`n"
    }
    
    $accordionInnerHtml = $accordionInnerHtml.Substring(0, $accordionInnerHtml.Length)

    $accordionInnerHtml += "</br>"
    $accordionInnerHtml += "`n"
    
    $accordionInnerHtml += "</br>"
    $accordionInnerHtml += "`n"

    $accordionInnerHtml += "</br>"
    $accordionInnerHtml += "`n"

    $accordionDiv = WriteHtmlElement -elementName "div" -class "accordion" -innerHtml $accordionInnerHtml

    Write-Output $accordionDiv
}

Function WriteHtmlAccordionGroup
{
    param
    (
        [string[]]
        [parameter(mandatory=$true)]
        $accordionItemList,

        [string]
        [parameter(mandatory=$true)]
        $title
    )

    $accordionGroupInnerHtml = [string]::Empty

    $heading = WriteHtmlElement -elementName "h2" -innerHtml $title

    $accordionGroupInnerHtml += $heading
    $accordionGroupInnerHtml += "`n"

    foreach ($item in $accordionItemList)
    {
        $accordionGroupInnerHtml += $item
        $accordionGroupInnerHtml += "`n"
    }

    $accordionGroupInnerHtml = $accordionGroupInnerHtml.Substring(0, $accordionGroupInnerHtml.Length-1)

    $accordionGroup = WriteHtmlElement -elementName "div" -class "accordion-group" -innerHtml $accordionGroupInnerHtml

    Write-Output $accordionGroup
}

 Function WriteHtmlAccordionItemForTable
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $title,

        [string]
        [parameter(mandatory=$true)]
        $tableId,

        [string[]]
        [parameter(mandatory=$true)]
        $tableHeaderColumns,

        [hashtable]
        [parameter(mandatory=$true)]
        $object,

        [string]
        [parameter(mandatory=$true)]
        $objectType,

        [hashtable]
        [parameter(mandatory=$false)]
        $objectToCompare
    )

    $headingInnerHtml = [string]::Empty

    $titleDiv = WriteHtmlElement -elementName "div" -class "title" -innerHtml $title

    $headingInnerHtml += $iconDiv
    $headingInnerHtml += "`n"
    $headingInnerHtml += $titleDiv

    $heading = WriteHtmlElement -elementName "a" -href "#" -class "heading" -innerHtml $headingInnerHtml

    $table = $null

    if($objectToCompare)
    {
        $table = WriteHtmlTable -tableId $tableId -tableHeaderColumns $tableHeaderColumns -object $object -objectToCompare $objectToCompare -objectType $objectType
    }
    else
    {
        $table = WriteHtmlTable -tableId $tableId -tableHeaderColumns $tableHeaderColumns -object $object -objectType $objectType
    }

    $tableContentDiv = WriteHtmlElement -elementName "div" -class "tableContent" -innerHtml $table

    $contentDiv = WriteHtmlElement -elementName "div" -class "content" -innerHtml $tableContentDiv

    $accordionItemInnerHtml = $heading
    $accordionItemInnerHtml += "`n"
    $accordionItemInnerHtml += $contentDiv

    $accordionItem = WriteHtmlElement -elementName "div" -class "accordion-item" -innerHtml $accordionItemInnerHtml

    Write-Output $accordionItem
 }

Function WriteHtmlAccordionItemForParagraph
{
    param
    (
        [System.Collections.Generic.List[string]]
        [parameter(mandatory=$true)]
        $messageList,

        [string]
        [parameter(mandatory=$true)]
        $title
    )

    $headingInnerHtml = [string]::Empty

    $titleDiv = WriteHtmlElement -elementName "div" -class "title" -innerHtml $title

    $headingInnerHtml += $iconDiv
    $headingInnerHtml += "`n"
    $headingInnerHtml += $titleDiv

    $heading = WriteHtmlElement -elementName "a" -href "#" -class "heading" -innerHtml $headingInnerHtml

    $paragraphInnerHtml = [string]::Empty

    foreach ($message in $messageList)
    {
        $paragraphInnerHtml += $message
        $paragraphInnerHtml += "`n"
    }

    $paragraphInnerHtml = $paragraphInnerHtml.Substring(0, $paragraphInnerHtml.Length-1)

    $paragraph = WriteHtmlElement -elementName "p" -class "issueText" -innerHtml $paragraphInnerHtml

    $contentDiv = WriteHtmlElement -elementName "div" -class "content" -innerHtml $paragraph

    $accordionItemInnerHtml = $heading
    $accordionItemInnerHtml += "`n"
    $accordionItemInnerHtml += $contentDiv

    $accordionItem = WriteHtmlElement -elementName "div" -class "accordion-item" -innerHtml $accordionItemInnerHtml

    Write-Output $accordionItem
}

Function WriteHtmlTable
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $tableId,

        [string[]]
        [parameter(mandatory=$true)]
        $tableHeaderColumns,

        [hashtable]
        [parameter(mandatory=$true)]
        $object,

        [hashtable]
        [parameter(mandatory=$false)]
        $objectToCompare,

        [string]
        [parameter(mandatory=$true)]
        $objectType
    )

    $tableHeader = WriteHtmlTableHeader -tableHeaderColumns $tableHeaderColumns

    $tableBody = $null

    if($objectToCompare)
    {
        $tableBody = WriteHtmlTableBody -object $object -objectToCompare $objectToCompare -objectType $objectType
    }
    else
    {
        $tableBody = WriteHtmlTableBody -object $object -objectType $objectType
    }

    $tableInnerHtml = $tableHeader
    $tableInnerHtml += "`n"
    $tableInnerHtml += $tableBody

    $table = WriteHtmlElement -elementName "table" -id $tableId -innerHtml $tableInnerHtml

    Write-Output $table
}

Function WriteHtmlTableHeader
{
    param
    (
        [string[]]
        [parameter(mandatory=$true)]
        $tableHeaderColumns
    )

    $rowInnerHtml = [string]::Empty

    foreach ($column in $tableHeaderColumns)
    {
        $rowInnerHtml += WriteHtmlElement -elementName "th" -scope "col" -innerHtml $column
        $rowInnerHtml += "`n"
    }

    $rowInnerHtml = $rowInnerHtml.Substring(0, $rowInnerHtml.Length-1)

    $row = WriteHtmlElement -elementName "tr" -innerHtml $rowInnerHtml

    $tableHeader = WriteHtmlElement -elementName "thead" -innerHtml $row
    
    Write-Output $tableHeader
}

Function WriteHtmlTableBody
{
    param
    (
        [hashtable]
        [parameter(mandatory=$true)]
        $object,

        [hashtable]
        [parameter(mandatory=$false)]
        $objectToCompare,

        [string]
        [parameter(mandatory=$true)]
        $objectType
    )

    $warningStyle = WriteStyleElement("background")("#ffff00")

    $bodyInnerHtml = [string]::Empty

    $object.GetEnumerator() | Foreach-Object {
        $attributeName = $_.Key
        $attributeValues = $_.Value

        $attributeValues2 = $null

        $rowInnerHtml = [string]::Empty

        $rowInnerHtml += WriteHtmlElement -elementName "td" -innerHtml $attributeName
        $rowInnerHtml += "`n"

        if ($attributeValues.Count -eq 1)
        {
            $rowInnerHtml += WriteHtmlElement -elementName "td" -innerHtml $attributeValues
            $rowInnerHtml += "`n"

            $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $objectType
            $rowInnerHtml += "`n"

            $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $attributeValues -class "eaAttributeValue"
            $rowInnerHtml += "`n"
        }
        else
        {
            $rowInnerHtml += WriteHtmlElement -elementName "td" -innerHtml "-- Multi-Valued --"
            $rowInnerHtml += "`n"

            $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $objectType
            $rowInnerHtml += "`n"

            foreach ($attributeValue in $attributeValues)
            {
                $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $attributeValue -class "eaAttributeValue"
                $rowInnerHtml += "`n"
            }
        }

        if($objectToCompare)
        {
            $attributeValues2 = $objectToCompare[$attributeName]

            if ($attributeValues2.Count -eq 0)
            {
                $rowInnerHtml += WriteHtmlElement -elementName "td" -innerHtml "-- No Value Retrieved --"
                $rowInnerHtml += "`n"

                $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $objectType
                $rowInnerHtml += "`n"

                $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value "-- No Value Retrieved --" -class "caAttributeValue"
                $rowInnerHtml += "`n"
            }
            elseif ($attributeValues2.Count -eq 1)
            {
                $rowInnerHtml += WriteHtmlElement -elementName "td" -innerHtml $attributeValues2
                $rowInnerHtml += "`n"

                $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $objectType
                $rowInnerHtml += "`n"

                $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $attributeValues2 -class "caAttributeValue"
                $rowInnerHtml += "`n"
            }
            else
            {
                $rowInnerHtml += WriteHtmlElement -elementName "td" -innerHtml "-- Multi-Valued --"
                $rowInnerHtml += "`n"

                $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $objectType
                $rowInnerHtml += "`n"

                foreach ($attributeValue in $attributeValues2)
                {
                    $rowInnerHtml += WriteHtmlElement -elementName "input" -type "hidden" -value $attributeValue -class "caAttributeValue"
                    $rowInnerHtml += "`n"
                }
            }
        }

        $rowInnerHtml = $rowInnerHtml.Substring(0, $rowInnerHtml.Length-1)

        if(($objectToCompare) -and (($attributeValues2.Count -eq 0) -or (Compare-Object $attributeValues $attributeValues2)))
        {
            $bodyInnerHtml += WriteHtmlElement -elementName "tr" -innerHtml $rowInnerHtml -style $warningStyle
        }
        else
        {
            $bodyInnerHtml += WriteHtmlElement -elementName "tr" -innerHtml $rowInnerHtml
        }

        $bodyInnerHtml += "`n"
    }

    $bodyInnerHtml = $bodyInnerHtml.Substring(0, $bodyInnerHtml.Length-1);

    $tableBody = WriteHtmlElement -elementName "tbody" -innerHtml $bodyInnerHtml

    Write-Output $tableBody
}

Function WriteHtmlMessage
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $message,

        [string]
        [parameter(mandatory=$false)]
        $color,

        [string]
        [parameter(mandatory=$false)]
        $fontSize,

        [string]
        [parameter(mandatory=$false)]
        $fontWeight,

        [string]
        [parameter(mandatory=$false)]
        $paddingLeft,

        [int]
        [parameter(mandatory=$true)]
        $numberOfLineBreaks
    )

    $style = [string]::Empty

    if (![string]::IsNullOrEmpty($color))
    {
        $style += WriteStyleElement("color")($color)
    }

    if (![string]::IsNullOrEmpty($fontSize))
    {
        $style += WriteStyleElement("font-size")($fontSize)
    }

    if (![string]::IsNullOrEmpty($fontWeight))
    {
        $style += WriteStyleElement("font-weight")($fontWeight)
    }

    if (![string]::IsNullOrEmpty($paddingLeft))
    {
        $style += WriteStyleElement("padding-left")($paddingLeft)
    }

    if (![string]::IsNullOrEmpty($style))
    {
        $style = $style.Substring(0, $style.Length-1)
    }
    
    $htmlMessage = WriteHtmlElement -elementName "span" -style $style -innerHtml $message

    for ($i = 0; $i -lt $numberOfLineBreaks; $i++)
    {
        $htmlMessage += "`n"
        $htmlMessage += "<br/>"
    }

    Write-Output $htmlMessage
}

Function WriteHtmlElement
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $elementName,

        [string]
        [parameter (mandatory=$false)]
        $id,

        [string]
        [parameter (mandatory=$false)]
        $class,

        [string]
        [parameter(mandatory=$false)]
        $scope,

        [string]
        [parameter(mandatory=$false)]
        $type,

        [string]
        [parameter(mandatory=$false)]
        $value,

        [string]
        [parameter(mandatory=$false)]
        $href,

        [string]
        [parameter(mandatory=$false)]
        $src,

        [string]
        [parameter(mandatory=$false)]
        $rel,

        [string]
        [parameter(mandatory=$false)]
        $style,

        [string]
        [parameter(mandatory=$false)]
        $onload,

        [string]
        [parameter(mandatory=$false)]
        $onclick,

        [string]
        [parameter(mandatory=$false)]
        $innerHtml
    )

    if ([string]::IsNullOrEmpty($elementName))
    {
    
    }

    $htmlElement = "<"
    $htmlElement += $elementName
    $htmlElement += " "
    
    if (![string]::IsNullOrEmpty($id))
    {
        $htmlElement += WriteHtmlAttribute("id")($id)
    }
    
    if (![string]::IsNullOrEmpty($class))
    {
        $htmlElement += WriteHtmlAttribute("class")($class)
    }
    
    if (![string]::IsNullOrEmpty($scope))
    {
        $htmlElement += WriteHtmlAttribute("scope")($scope)
    }

    if (![string]::IsNullOrEmpty($type))
    {
        $htmlElement += WriteHtmlAttribute("type")($type)
    }

    if (![string]::IsNullOrEmpty($value))
    {
        $htmlElement += WriteHtmlAttribute("value")($value)
    }

    if (![string]::IsNullOrEmpty($href))
    {
        $htmlElement += WriteHtmlAttribute("href")($href)
    }
    
    if (![string]::IsNullOrEmpty($src))
    {
        $htmlElement += WriteHtmlAttribute("src")($src)
    }

    if (![string]::IsNullOrEmpty($rel))
    {
         $htmlElement += WriteHtmlAttribute("rel")($rel)
    }

    if (![string]::IsNullOrEmpty($style))
    {
        $htmlElement += WriteHtmlAttribute("style")($style)
    }

    if (![string]::IsNullOrEmpty($onload))
    {
        $htmlElement += WriteHtmlAttribute("onload")($onload)
    }
    
    if (![string]::IsNullOrEmpty($onclick))
    {
        $htmlElement += WriteHtmlAttribute("onclick")($onclick)
    }
    
    # Remove last space character
    $htmlElement = $htmlElement.Substring(0, $htmlElement.Length-1)

    $htmlElement += ">"
    
    if (![string]::IsNullOrEmpty($innerHtml))
    {
        $htmlElement += "`r`n"
        
        $innerHtmlLines = $innerHtml.Split("`n")

        foreach ($line in $innerHtmlLines)
        {
            $htmlElement += "    "
            $htmlElement += $line
            $htmlElement += "`n"
        }
    }

    $htmlElement += "</"
    $htmlElement += $elementName
    $htmlElement += ">"

    Write-Output $htmlElement
}

Function WriteHtmlAttribute
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $attributeName,

        [string]
        [parameter(mandatory=$true)]
        $attributeValue
    )

    $htmlAttribute = $attributeName
    $htmlAttribute += "="
    $htmlAttribute += """"
    $htmlAttribute += $attributeValue
    $htmlAttribute += """"
    $htmlAttribute += " "

    Write-Output $htmlAttribute
}

Function WriteStyleElement
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $styleElementName,

        [string]
        [parameter(mandatory=$true)]
        $styleElementValue
    )

    $styleElement = [string]::Empty

    $styleElement += $styleElementName
    $styleElement += ": "
    $styleElement += $styleElementValue
    $styleElement += "; "

    Write-Output $styleElement
}

Function WriteHyperlink
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $url,

        [string]
        [parameter(mandatory=$true)]
        $text
    )

    $hyperlink = WriteHtmlElement -elementName "a" -href $url -style "color:#0078d7" -innerHtml $text

    Write-Output $hyperlink
}

Function WriteAttributePopupWindow
{
    $attributePopupWindow = @"
<div class="popupContainer" id="singleValuePopupContainerId">	
    <div class="attributeWindow"> 
        <div class="ObjectTypeHeading">
            <span class="close" onclick="closePopUpWindow()">&times;</span>
            <p class="ObjectTypeText" id="singleValueObjectTypeTextId"></p>
        </div>
        <div class="AttributeNameField">
            <p class="AttributeNameFieldText">Attribute Name:</p>
        </div>
        <div class="AttributeName">
            <p class="AttributeNameText" id="singleValueAttributeNameTextId"></p>
        </div>
        <div class="AttributeValueField">
            <p class="AttributeValueFieldText">Attribute Value:</p>
        </div>
        <div id="SingleValueAttributeValue">
            <p class="AttributeValueText" id="singleValueAttributeValueTextId"></p>
        </div>
    </div>
</div>

<div class="popupContainer" id="multiValuePopupContainerId">	
    <div class="attributeWindow"> 
        <div class="ObjectTypeHeading">
            <span class="close" onclick="closePopUpWindow()">&times;</span>
            <p class="ObjectTypeText" id="multiValueObjectTypeTextId"></p>
        </div>
        <div class="AttributeNameField">
            <p class="AttributeNameFieldText">Attribute Name:</p>
        </div>
        <div class="AttributeName">
            <p class="AttributeNameText" id="multiValueAttributeNameTextId"></p>
        </div>
        <div class="AttributeValueField">
            <p class="AttributeValueFieldText">Attribute Values:</p>
        </div>
        <div id="MultiValueAttributeValue">
            <table id="multiValueAttributeValueTableId">
            </table>
        </div>
    </div>
</div>

<div class="popupContainer" id="comparisonPopupContainerId">	
    <div class="attributeWindow"> 
        <div class="ObjectTypeHeading">
            <span class="close" onclick="closePopUpWindow()">&times;</span>
            <p class="ObjectTypeText" id="comparisonObjectTypeTextId"></p>
        </div>
        <div class="AttributeNameField">
            <p class="AttributeNameFieldText">Attribute Name:</p>
        </div>
        <div class="AttributeName">
            <p class="AttributeNameText" id="comparisonAttributeNameTextId"></p>
        </div>
        <div class="AttributeValueField">
            <p class="AttributeValueFieldText">Attribute Value(s) retrieved by provided account:</p>
        </div>
        <div id="MultiValueAttributeValue">
            <table id="comparisonAttributeValueTableOneId">
            </table>
        </div>
        <div class="AttributeValueField">
            <p class="AttributeValueFieldText">Attribute Value(s) retrieved by Connector:</p>
        </div>
        <div id="MultiValueAttributeValue">
            <table id="comparisonAttributeValueTableTwoId">
            </table>
        </div>
    </div>
</div>

"@

    Write-Output $attributePopupWindow
}

Function WriteHtmlStyle
{
    $htmlStyle = @"
html 
{
    font-family: "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif;
}

body
{
    background: #FFFFFF;
    margin: 0;
}

.banner
{
    padding-top: 10px;
    padding-bottom: 10px;
    padding-left: 30px;
    background: #000000;
}

.header 
{
    text-align: center;
    background: #252525;
    padding-bottom: 5px;
}

h1 
{
    color: #000000;
}

h2 
{
    border-bottom: solid 2px #CCCCCC;
    padding-bottom: 3px;
    color: #252525;
    font-weight: 500;
}

.accordion 
{
    width: 100%;
    max-width: 95rem;
    margin: 0 auto;
    padding: 2rem;
}

.accordion-group 
{
    position: relative;
}

.accordion-item 
{
    position: relative;
}
.accordion-item.active .heading 
{
    color: #0078d7;
}

.accordion-item.active .icon 
{
    background: #fefcff;
}

.accordion-item.active .icon:before 
{
    background: #0078d7;
}

.accordion-item.active .icon:after 
{
    width: 0;
}

.accordion-item .heading 
{
    display: block;
    text-decoration: none;
    color: #0078d7;
    font-weight: 500;
    font-size: 1rem;
    position: relative;
    padding: 1.5rem 0 1.5rem;
    -webkit-transition: 0.3s ease-in-out;
    transition: 0.3s ease-in-out;
}

@media (min-width: 40rem) 
{
    .accordion-item .heading 
    {
        font-size: 1.2rem;
    }
}

.accordion-item .heading:hover
{
    text-decoration: underline;
}

.accordion-item .heading:hover .icon
{
    background: #fefcff;
}

.accordion-item .heading:hover .icon:before, .accordion-item .heading:hover .icon:after
{
    background: #1F45FC;
}

.accordion-item .icon 
{
    display: block;
    position: absolute;
    top: 50%;
    left: 0;
    width: 3rem;
    height: 3rem;
    border: 2px solid #000080;
    border-radius: 3px;
    -webkit-transform: translateY(-50%);
            transform: translateY(-50%);
}

.accordion-item .icon:before, .accordion-item .icon:after 
{
    content: '';
    width: 1.25rem;
    height: 0.25rem;
    background: #000080;
    position: absolute;
    border-radius: 3px;
    left: 50%;
    top: 50%;
    -webkit-transition: 0.3s ease-in-out;
    transition: 0.3s ease-in-out;
    -webkit-transform: translate(-50%, -50%);
            transform: translate(-50%, -50%);
}

.accordion-item .icon:after 
{
    -webkit-transform: translate(-50%, -50%) rotate(90deg);
            transform: translate(-50%, -50%) rotate(90deg);
    z-index: -1;
}

.accordion-item .content 
{
    width: 100%;
    display: none;
    background: #ffffff;
}

.accordion-item .content
{
    margin-top: 0;
}


@media (min-width: 40rem)
{
    .accordion-item .content
    {
        line-height: 1.75;
    }
}

.accordion-item .content .tableContent
{
    max-height: 600px;
    overflow-y: auto;
}

table
{
    width: 90%;
    border-collapse: collapse;
    margin-left: auto;
    margin-right: auto;
}

th 
{ 
    background: #ffffff; 
    color: #252525; 
    font-weight: bold; 
    text-align: left;
    padding-left: 50px;
    padding-top: 15px;
    padding-bottom: 15px;
    border-bottom: 1px solid #A9A9A9;
}

tr
{
    background: #ffffff;
}

tr:hover 
{
    background: #b4d2e5;
}

td
{
    width: 50%;
    color: #252525;
    text-align: left;
    padding-left: 50px;
    padding-top: 15px;
    padding-bottom: 15px;
    border-bottom: 1px solid #A9A9A9;
}

.popupContainer
{
    display: none;
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    z-index: 1;
    background-color: rgba(0, 0, 0, 0.6);
}

.attributeWindow
{
    width: 40%;
    background: #d9d9d9;
    margin-left: auto;
    margin-right: auto;
    padding-bottom: 20px;
    opacity: 1.0;
    vertical-align: middle;
    margin-top: 100px;
}

.ObjectTypeHeading
{
    width: 100%;
    background: #003c6c;
}

.ObjectTypeText
{
    padding-left: 5%;
    padding-bottom: 10px;
    padding-top: 10px;
    color: #fff;
    font-weight: bold;
    font-size: 18px;
}

.AttributeNameField
{
    width: 100%;
}

.AttributeNameFieldText
{
    padding-left: 5%;
    color: #000;
    font-weight: bold;
    font-size: 16x;
}

.AttributeName
{	
    width: 90%;
    background: #fff;
    margin-left: auto;
    margin-right: auto;
}

.AttributeNameText
{
    padding-left: 4px;
    padding-top: 5px;
    padding-bottom: 5px;
    font-size: 16x;
    color: #000;
}

.AttributeValueField
{
    width: 100%;
    padding-top: 5px;
}

.AttributeValueFieldText
{
    padding-left: 5%;
    color: #000;
    font-weight: bold;
    font-size: 16x;
}

#SingleValueAttributeValue
{
    width: 90%;
    background: #fff;
    margin-left: auto;
    margin-right: auto;
}

#MultiValueAttributeValue
{
    width: 90%;
    background: #fff;
    margin-left: auto;
    margin-right: auto;
    max-height: 200px;
    overflow-y: auto;
}

.AttributeValueText
{
    padding-left: 4px;
    padding-top: 5px;
    padding-bottom: 5px;
    font-size: 16x;
    color: #000;
}

.close
{
    color: #000000;
    float: right;
    font-size: 32px;
    font-weight: bold;
    padding-right: 10px;
}

.close:hover
{
    color: #ff0000;
}

#ObjectDnId
{
    font-size: 20px;
}

.issueText
{
    font-size: 14px;
}
"@

    $htmlStyleElement = WriteHtmlElement -elementName "style" -innerHtml $htmlStyle

    Write-Output $htmlStyleElement
}

Function WriteHtmlScript
{
    $htmlScript = @"
`$('.accordion-item .heading').on('click', function(e) {
    e.preventDefault();

    // Add the correct active class
    if(`$(this).closest('.accordion-item').hasClass('active')) {
        // Remove active classes
        `$('.accordion-item').removeClass('active');
    } else {
        // Remove active classes
        `$('.accordion-item').removeClass('active');

        // Add the active class
        `$(this).closest('.accordion-item').addClass('active');
    }

    // Show the content
    var `$content = `$(this).next();
    `$content.slideToggle(200);
    `$('.accordion-item .content').not(`$content).slideUp('fast');
    
    `$('html, body').animate({scrollTop: `$content.offset().top-100}, 600);
});

function addRowHandlers()
{	
    addObjectTableRowHandler("ADObjectTable");
    addObjectTableRowHandler("AADConnectObjectTable");
    addObjectTableRowHandler("AADObjectTable");
    addObjectComparisonTableRowHandler("ADObjectAttributeComparisonTable");
};

function addObjectTableRowHandler(tableId)
{
    var objectTable = document.getElementById(tableId);

    if(objectTable != null)
    {
        var rows = objectTable.getElementsByTagName("tr");
    
        for (i = 1; i < rows.length; i++)
        {
            var currentRow = rows[i];
        
            var createRowClickHandler =
                function(row)
                {
                    return function(){
                        var cols = row.getElementsByTagName("td");
                        var values = row.getElementsByTagName("input");
                    
                        if (values.length == 2)
                        {
                            document.getElementById("singleValueObjectTypeTextId").innerHTML = values[0].value + " Details";
                            document.getElementById("singleValueAttributeNameTextId").innerHTML = cols[0].innerHTML;
                            document.getElementById("singleValueAttributeValueTextId").innerHTML = values[1].value;
                            document.getElementById("singleValuePopupContainerId").style.display = "block";
                        }
                        else
                        {
                            document.getElementById("multiValueObjectTypeTextId").innerHTML = values[0].value + " Details";
                            document.getElementById("multiValueAttributeNameTextId").innerHTML = cols[0].innerHTML;
                        
                            var multiValueTable = document.getElementById("multiValueAttributeValueTableId");
                        
                            for (k = 1; k < values.length; k++)
                            {
                                var multiValueRow = multiValueTable.insertRow(k-1);
                                var multiValueCell = multiValueRow.insertCell(0);
                            
                                multiValueCell.innerHTML = values[k].value;
                            }
                        
                            document.getElementById("multiValuePopupContainerId").style.display = "block";
                        }
                    };
                };
        
            currentRow.onclick = createRowClickHandler(currentRow);
        }
    }
}

function addObjectComparisonTableRowHandler(tableId)
{
    var objectTable = document.getElementById(tableId);

    if(objectTable != null)
    {
        var rows = objectTable.getElementsByTagName("tr");
    
        for (i = 1; i < rows.length; i++)
        {
            var currentRow = rows[i];
        
            var createRowClickHandler =
                function(row)
                {
                    return function(){
                        var cols = row.getElementsByTagName("td");
                        var values = row.getElementsByTagName("input");

                        var values1 = row.getElementsByClassName("eaAttributeValue");
                        var values2 = row.getElementsByClassName("caAttributeValue");
                    
                        document.getElementById("comparisonObjectTypeTextId").innerHTML = values[0].value + " Details";
                        document.getElementById("comparisonAttributeNameTextId").innerHTML = cols[0].innerHTML;
                        
                        var multiValueTable = document.getElementById("comparisonAttributeValueTableOneId");
                        
                        for (k = 0; k < values1.length; k++)
                        {
                            var multiValueRow = multiValueTable.insertRow(k-1);
                            var multiValueCell = multiValueRow.insertCell(0);
                            
                            multiValueCell.innerHTML = values1[k].value;
                        }

                        var multiValueTable2 = document.getElementById("comparisonAttributeValueTableTwoId");
                        
                        for (k = 0; k < values2.length; k++)
                        {
                            var multiValueRow = multiValueTable2.insertRow(k-1);
                            var multiValueCell = multiValueRow.insertCell(0);
                            
                            multiValueCell.innerHTML = values2[k].value;
                        }
                        
                        document.getElementById("comparisonPopupContainerId").style.display = "block";
                    };
                };
        
            currentRow.onclick = createRowClickHandler(currentRow);
        }
    }
}

function closePopUpWindow()
{
    document.getElementById("multiValueAttributeValueTableId").innerHTML = "";
    document.getElementById("comparisonAttributeValueTableOneId").innerHTML = "";
    document.getElementById("comparisonAttributeValueTableTwoId").innerHTML = "";
    document.getElementById("singleValuePopupContainerId").style.display = "none";
    document.getElementById("multiValuePopupContainerId").style.display = "none";
    document.getElementById("comparisonPopupContainerId").style.display = "none";
}
"@

    $htmlScriptElement = WriteHtmlElement -elementName "script" -innerHtml $htmlScript

    Write-Output $htmlScriptElement
}
# SIG # Begin signature block
# MIIoNgYJKoZIhvcNAQcCoIIoJzCCKCMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCTjNecArQtekvp
# fg+jfXT8byEgXBLTIplue5l4J1DP4qCCDYIwggYAMIID6KADAgECAhMzAAADXJXz
# SFtKBGrPAAAAAANcMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwNDA2MTgyOTIyWhcNMjQwNDAyMTgyOTIyWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDijA1UCC84R0x+9Vr/vQhPNbfvIOBFfymE+kuP+nho3ixnjyv6vdnUpgmm6RT/
# pL9cXL27zmgVMw7ivmLjR5dIm6qlovdrc5QRrkewnuQHnvhVnLm+pLyIiWp6Tow3
# ZrkoiVdip47m+pOBYlw/vrkb8Pju4XdA48U8okWmqTId2CbZTd8yZbwdHb8lPviE
# NMKzQ2bAjytWVEp3y74xc8E4P6hdBRynKGF6vvS6sGB9tBrvu4n9mn7M99rp//7k
# ku5t/q3bbMjg/6L6mDePok6Ipb22+9Fzpq5sy+CkJmvCNGPo9U8fA152JPrt14uJ
# ffVvbY5i9jrGQTfV+UAQ8ncPAgMBAAGjggF/MIIBezArBgNVHSUEJDAiBgorBgEE
# AYI3TBMBBgorBgEEAYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUXgIsrR+tkOQ8
# 10ekOnvvfQDgTHAwRQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEWMBQGA1UEBRMNMjMzMTEwKzUwMDg2ODAfBgNVHSMEGDAWgBRI
# bmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEt
# MDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBABIm
# T2UTYlls5t6i5kWaqI7sEfIKgNquF8Ex9yMEz+QMmc2FjaIF/HQQdpJZaEtDM1Xm
# 07VD4JvNJEplZ91A4SIxjHzqgLegfkyc384P7Nn+SJL3XK2FK+VAFxdvZNXcrkt2
# WoAtKo0PclJOmHheHImWSqfCxRispYkKT9w7J/84fidQxSj83NPqoCfUmcy3bWKY
# jRZ6PPDXlXERRvl825dXOfmCKGYJXHKyOEcU8/6djs7TDyK0eH9ss4G9mjPnVZzq
# Gi/qxxtbddZtkREDd0Acdj947/BTwsYLuQPz7SNNUAmlZOvWALPU7OOVQlEZzO8u
# Ec+QH24nep/yhKvFYp4sHtxUKm1ZPV4xdArhzxJGo48Be74kxL7q2AlTyValLV98
# u3FY07rNo4Xg9PMHC6sEAb0tSplojOHFtGtNb0r+sioSttvd8IyaMSfCPwhUxp+B
# Td0exzQ1KnRSBOZpxZ8h0HmOlMJOInwFqrCvn5IjrSdjxKa/PzOTFPIYAfMZ4hJn
# uKu15EUuv/f0Tmgrlfw+cC0HCz/5WnpWiFso2IPHZyfdbbOXO2EZ9gzB1wmNkbBz
# hj8hFyImnycY+94Eo2GLavVTtgBiCcG1ILyQabKDbL7Vh/OearAxcRAmcuVAha07
# WiQx2aLghOSaZzKFOx44LmwUxRuaJ4vO/PRZ7EzAMIIHejCCBWKgAwIBAgIKYQ6Q
# 0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
# dGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5
# WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQD
# Ex9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4
# BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe
# 0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato
# 88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v
# ++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDst
# rjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN
# 91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4ji
# JV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmh
# D+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbi
# wZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8Hh
# hUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaI
# jAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTl
# UAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNV
# HQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQF
# TuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29m
# dC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNf
# MjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNf
# MjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcC
# ARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnlj
# cHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5
# AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oal
# mOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0ep
# o/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1
# HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtY
# SWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInW
# H8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZ
# iWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMd
# YzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7f
# QccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKf
# enoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOpp
# O6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZO
# SEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jv
# c29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANclfNIW0oEas8AAAAAA1ww
# DQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINjsqvfH
# Jd55trC0CgwXBxyPYGeIMiGmbQ2KaPjg9qDRMEIGCisGAQQBgjcCAQwxNDAyoBSA
# EgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20w
# DQYJKoZIhvcNAQEBBQAEggEAhZxL6lAhbVfh9S0OkncFG8j4TCFlcG7nqoKnNFBR
# vSlwgtPeBgKi8Gm+W7AtGWaYV1Djnuno/c2Ca0O5Y4HorJi9QIJzMt69MYR5OoLu
# X3nUvYG8ewJrC4J5KqIceRjdr0WVYgJmBO155JQoWfvyZoDB5Isnt+bvHDwy3HfU
# kyVO3O/aj2MAoloZVW+wT+hYvhK3MI0YRkO/xrOpeoOiy1MOi/ydLLZ6QOBRi1su
# th26n8GcJ+FdyKImDUaaG0ZiF6lhORPlouRyF0EjALWJCh2oglHow7BYqu00Hjx4
# o+8hElYbiSvbl6j09y2qlIEz8vPv8BasCQSGdV4ZXmHBRqGCF5QwgheQBgorBgEE
# AYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUD
# BAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoD
# ATAxMA0GCWCGSAFlAwQCAQUABCA02SpYFNcCWOQQ71xGiuX7I3GheykJemOnH/eV
# AxAmIQIGZQtfrE2yGBMyMDIzMTAwNDE5MzA1MS4xMzVaMASAAgH0oIHRpIHOMIHL
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxN
# aWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRT
# UyBFU046MzMwMy0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAcyGpdw369lhLQABAAAB
# zDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAe
# Fw0yMzA1MjUxOTEyMDFaFw0yNDAyMDExOTEyMDFaMIHLMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmlj
# YSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046MzMwMy0wNUUw
# LUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDMsSIF8e9NmEc+83NVZGgWWZi/
# wBYt8zhxAfSGM7xw7K7CbA/1A4GhovPvkIY873tnyzdyZe+6YHXx+Rd618lQDmmm
# 5X4euiYG53Ld7WIK+Dd+hyi0H97D6HM4ZzGqovmwB0fZ3lh+phJLoPT+9yrTLFzk
# kKw2Vcb7wXMBziD0MVVYbmwRlRaypTntl39IENCEijW9j6MElTyXP2zrc0OthQN5
# RrMTY5iZja3MyHCFmYMGinmHftsaG3Ydi8Ga8BQjdtoTm5dVhnqs2qKNEOqZSon2
# 8R4Xff0tlJL5UHyI3bywH/+zQeJu8qnsSCi8VFPOsZEb6cZzhXHaAiSGtdKAbQRa
# AIhExbIUpeJypC7l+wqKC3BO9ADGupB9ZgUFbSv5ECFjMDzbfm8M5zz2A4xYNPQX
# qZv0wGWL+jTvb7kFYiDPPe+zRyBbzmrSpObB7XqjqzUFNKlwp+Mx15k1F7FMs5EM
# 2uG68IQsdAGBkZbSDmuGmjPbZ7dtim+XHuh3NS6JmXYPS7rikpCbUsMZMn5eWxiW
# FIk6f00skR4RLWmh0N6Oq+KYI1fA59LzGiAbOrcxgvQkRo3OD4o1JW9z1TNMwEbk
# zPrXMo8rrGsuGoyYWcsm9xhd0GXIRHHC64nzbI3e0G5jqEsWQc4uaQeSRyr70KRi
# jzVyWjjYfsEtvVMlJwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFIKmHGRdPIdLRXts
# R5XRSyM3+2kMMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1Ud
# HwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3Js
# L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggr
# BgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIw
# MTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgw
# DgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQB5GUMo9XviUl3g72u8
# oQTorIKDoAdgWZ4LQ9+dAEQCmaetsThkxbNm15seu7GmwpZdhMQN8TNddGki5s5I
# e+aA2VEo9vZz31llusHBXAVrQtpufQqtIA+2nnusfaYviitr6p5kVT609LITOYgd
# KRWEpfx/4yT5R9yMeKxoxkk8tyGiGPZK40ST5Z14OPdJfVbkYeCvlLQclsX1+WBZ
# Nx/XZvazJmXjvYjTuG0QbZpxw4ZO3ZoffQYxZYRzn0z41U7MDFlXo2ihfasdbHuu
# a6kpHxJ9AIoUevh3mzvUxYp0u0z3wYDPpLuo+M2VYh8XOCUB0u75xG3S5+98TKmF
# bqZYgpgr6P+YKeao2YpB1izs850YSzuwaX7kRxAURlmN/j5Hv4wabnOfZb36mDqJ
# p4IeGmwPtwI8tEPsuRAmyreejyhkZV7dfgJ4N83QBhpHVZlB4FmlJR8yF3aB15QW
# 6tw4CaH+PMIDud6GeOJO4cQE+lTc6rIJmN4cfi2TTG7e49TvhCXfBS2pzOyb9Yem
# Sm0krk8jJh6zgeGqztk7zewfE+3shQRc74sXLY58pvVoznfgfGvy1llbq4Oey96K
# ouwiuhDtxuKlTnW7pw7xaNPhIMsOxW8dpSp915FtKfOqKR/dfJOsbHDSJY/iiJz4
# mWKAGoydeLM6zLmohRCPWk/Q5jCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkA
# AAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRl
# IEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVow
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX
# 9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1q
# UoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8d
# q6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byN
# pOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2k
# rnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4d
# Pf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgS
# Uei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8
# QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6Cm
# gyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzF
# ER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQID
# AQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU
# KqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1
# GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0
# bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMA
# QTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbL
# j+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1p
# Y3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIz
# LmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwU
# tj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN
# 3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU
# 5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5
# KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGy
# qVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB6
# 2FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltE
# AY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFp
# AUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcd
# FYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRb
# atGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQd
# VTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2Eg
# T3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjMzMDMtMDVFMC1E
# OTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEw
# BwYFKw4DAhoDFQBOTuZ3uYfiihS4zRToxisDt9mJpKCBgzCBgKR+MHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6MeqQzAiGA8y
# MDIzMTAwNDA5MDkyM1oYDzIwMjMxMDA1MDkwOTIzWjB0MDoGCisGAQQBhFkKBAEx
# LDAqMAoCBQDox6pDAgEAMAcCAQACAickMAcCAQACAhPvMAoCBQDoyPvDAgEAMDYG
# CisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEA
# AgMBhqAwDQYJKoZIhvcNAQELBQADggEBALGP4DiN7ida2cBAd/2MhcdApOoTetnY
# ya2m+1GM3EP59Dy464aRdPGmWTeGbFnd3+imdUUijJd2bNCGMdkecoD81qqFm5+k
# 3p/h/sDy54tbiVvVwl6rTtsh3wBZrJvRrn2nwcOaGecDV2PJ1kbxUaNA4FrbqUoV
# lcLNyrsUoL1Yk6ry3wpchpw70EsD/0vfFEA1EP0NWUpL1MmM3eQw8lqAOW6HMetm
# ZtRT/qXGdHiOuxLGyVQPFxj2PNsB06d2AjcOL+GTXl4X5jYf8fpJ2VcD3hhpY5l/
# mquTbGasjTFKHwlMnDn3XMzlzNnrGyV1+44taR3/qs/igHQOTpEqdM4xggQNMIIE
# CQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYw
# JAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAcyGpdw3
# 69lhLQABAAABzDANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqG
# SIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCC5tZMfcHRGsEzg6VjLq9kyylELDhtJ
# G7LXrMmze+8glDCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EINbuZQHC/sMC
# E+cgKVSkwKpDICfnefFZYgDbF4HrcFjmMIGYMIGApH4wfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAHMhqXcN+vZYS0AAQAAAcwwIgQg81i1yTHfGKw0
# qMNqKk83/R6lblUkfjg6lyxCSQ9/Z1owDQYJKoZIhvcNAQELBQAEggIARrhyfkQD
# yEIfQklYFPSZXns8LMeMX7D3f8C2rTDPCOKISC7uMnjuudeZlJi6P+BxEFudQAnl
# QazVEvfGLr3ZMPooWFIHS7pljuGlqur+r+MmTS1OpGtRd428NFpp7Nkekc9owNT/
# nXmuELjmZ4P5l9RQslftE6IMRw6ffNpo8dSiXdIelqhrizqkinWJ54zANTVnUQLk
# f99FZHcrAmpzsURfpDe0FbBeDHJrwVtwzHxHWtyS/E9FrpacYoVud/gRWHpPrNPZ
# Qt6zMsBGg3GjNiE4jviXZaze6GsbXGVtVhNiAHQuDp94CqnxuNU7Es4BCItQ7cdv
# r/5FysA8XNwy5+YDVYGqBiO/nR8A123nWO4ZyYYNGRaZlOEInWQ35+OJHes23ia5
# w3V2bIr03U7Fh+D3BJmwmW+o6wJBWoM227la7NtHa7kvjK8nRJLy5G08g7nL3XWG
# zeXMrdCwMJJfH+TWWWM+rpAWIVSD4GyqwjTgEBwLLua25qeknZaK0pKJh/BNFYg1
# J3aSYMUssfBMgy+C5nPLZEl/g5YEiHTSx8WY0IhuMPIp13PfjofXW6X600nY9/Jy
# BpnXF297on1If6UddM+3aJOyyIx1802h4ozVgnw+JFe6y7OO/XU62qmftNr5Z9Od
# yUJNk4fohiA8RFo7R9Gpzhy8Phy0RFA7EtA=
# SIG # End signature block
