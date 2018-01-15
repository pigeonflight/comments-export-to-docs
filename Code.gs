function onOpen() {
  DocumentApp.getUi().createAddonMenu().addItem('Comments Export', 'showSidebar').addToUi();  
}

function showSidebar() {
 var html = HtmlService.createTemplateFromFile("Comments Exporter").evaluate()
   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle("Comments Exporter");
  DocumentApp.getUi().showSidebar(html)
}

function getComments(){
var document_id = DocumentApp.getActiveDocument().getId();
return Drive.Comments.list(document_id,{
                                         maxResults: 100
                                        }).items;
  
}



function createDocument(){
 // Create and open a document
var current_doc = DocumentApp.getActiveDocument();
new_title = "Comments On: "
            + current_doc.getName()
var doc = DocumentApp.create(new_title);
return doc  
}

function extractContents(item){
    return item.contents
}

function decodeHtml(html) {
  text = XmlService.createText(html)
  return text.getText()
}

function populateDocument(doc){
var context_style = {};
context_style[DocumentApp.Attribute.FONT_FAMILY] = 'Montserrat';
context_style[DocumentApp.Attribute.FONT_SIZE] = 10;
context_style[DocumentApp.Attribute.ITALIC] = true;
var default_style = {};
default_style[DocumentApp.Attribute.FONT_FAMILY] = 'Montserrat';
default_style[DocumentApp.Attribute.ITALIC] = false;
default_style[DocumentApp.Attribute.FONT_SIZE] = 12;
var body = doc.getBody();
  
var comments = getComments();
// Append comments as paragraphs
  para = body.appendParagraph("Comments");
  para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
// var comment_contents = comments.map(extractContents)
for(var i in comments)
{
  comment = JSON.parse(comments[i]);
  replies = comment.replies;
  var reply_ = "";
  
  if (typeof replies !== 'undefined' && replies.length > 0) {
    // the array is defined and has at least one element
    replies.forEach(function(reply){
    if (reply.content.indexOf('[') != -1 && reply.content.indexOf(']') > 4){
     comment_heading = body.appendParagraph(decodeHtml(reply.content));
     comment_heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else {
      reply_ = reply_ + "\n" + decodeHtml(reply.content);
    }
     });
    
  }
  
  
  comment_context = body.appendParagraph(comment.context.value);
  comment_context.setAttributes(context_style)
  comment_context.setSpacingAfter(1);
  
  comment_body = body.appendParagraph(decodeHtml(comment.content));
  comment_body.setAttributes(default_style)
  comment_body.setLineSpacing(1.5)
  
  comment_reply_ = body.appendParagraph(reply_);
  comment_reply_.setAttributes(default_style)
  comment_reply_.setLineSpacing(1.5)
  
}
//comment_contents.map(body.appendParagraph);
return doc
}
function generateCommentDocument(){
 var mydoc = createDocument();
 populateDocument(mydoc); 
 return mydoc.getUrl() 
}
