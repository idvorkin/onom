<Query Kind="Statements" />

var elementText = 
         "<a href=\"onenote:SmartTagStorage.one#Model%205844974a-21ac-4de5-83cc-9ec9aa362b0d&amp;section-id={D6DF7A74-EF20-476B-AC90-8D6D6BBBE4DC}&amp;page-id={2C9EA676-6657-4435-B4A1-E5AC6720AE34}&amp;end&amp;extraId={3EF0B7D7-1857-0633-04EE-212D80860CA2}{1}{E19543199418872411574120135711402651528027051}&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">#</a><a" +
         " href=\"onenote:Topics.one#hello&amp;section-id={25BAB527-01EA-417E-B55F-D52F8E5FB2EC}&amp;page-id={68DDB3B7-F9E5-455F-90B5-68E2823C5188}&amp;end&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">hello</a> world";

var elementText2= "<a"+
" href=\"onenote:SmartTagStorage.one#Model%20981eff41-0428-41f4-8f0c-3ce1ea3546e9&amp;section-id={D6DF7A74-EF20-476B-AC90-8D6D6BBBE4DC}&amp;page-id={EEB32202-CAF4-41B7-AF5A-4BB30DD4327E}&amp;end&amp;extraId={3EF0B7D7-1857-0633-04EE-212D80860CA2}{1}{E19468513483745559900720152138099127754830691}&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">#</a><a"+
" href=\"onenote:Topics.one#processedTag&amp;section-id={25BAB527-01EA-417E-B55F-D52F8E5FB2EC}&amp;page-id={A28B1AD1-4BC0-474A-ABB3-40575F5C6083}&amp;end&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">processedTag</a> is Processed";


var elementText3 =" <a\n" +
"href=\"onenote:SmartTagStorage.one#Model%20981eff41-0428-41f4-8f0c-3ce1ea3546e9&amp;section-id={D6DF7A74-EF20-476B-AC90-8D6D6BBBE4DC}&amp;page-id={EEB32202-CAF4-41B7-AF5A-4BB30DD4327E}&amp;end&amp;extraId={3EF0B7D7-1857-0633-04EE-212D80860CA2}{1}{E19468513483745559900720152138099127754830691}&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">#</a><a\n"+
"href=\"onenote:Topics.one#processedTag&amp;section-id={25BAB527-01EA-417E-B55F-D52F8E5FB2EC}&amp;page-id={A28B1AD1-4BC0-474A-ABB3-40575F5C6083}&amp;end&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">processedTag</a> is Processed\n";

var  augmentedSmartTagFullTextPattern = "extraId=(?<extraId>.{85,90}}).*>#</a>.*>(?<tagName>.*)</a> (?<fullText>.+)";
var augmenetedfullTextMatch = Regex.Match(elementText3,  augmentedSmartTagFullTextPattern,RegexOptions.Singleline);
augmenetedfullTextMatch.Success.Dump();
augmenetedfullTextMatch.Dump();
var augmentedfullText = "#" + augmenetedfullTextMatch.Groups["tagName"].Value + " "+ augmenetedfullTextMatch.Groups["fullText"];
augmentedfullText.Dump();
