// ======================================
// Displays HTML code for the current
// section (only for script debugging)
// ======================================

function Run() {
  var sel = document.getSelection();
  if (!sel) 
	return false;
  var pe = sel.anchorNode;
  while (pe && pe.tagName != 'DIV' )
    pe = pe.parentNode;

  if (pe)
	alert(pe.outerHTML);
}
