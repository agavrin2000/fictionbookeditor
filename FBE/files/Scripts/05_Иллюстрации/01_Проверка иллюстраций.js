//======================================
// ����� ������������� ��� �������������� �����������
//                  Alex Saveliev script for FBW2         version 1.1
//======================================

// TO DO: Go to the first broken image if any

function Run()
{
 var lists=$('#tiCover select');

 var covers=""; for(var i=0; i<lists.length; i++) if(lists[i].id=='href') covers+=lists[i].value+'|';

 // Find unused images

 var imgs = fbw_body.getElementsByTagName("IMG"), imgs_len=imgs.length;
 var bins= $('#binobj DIV'), bins_len=bins.length;

 var unused="";


 for(var i=0; i<bins_len; i++)
 {
  var found=false; var bin="fbw-internal:#"+bins[i].id.value;

  for(var j=0; j<imgs_len; j++) if(imgs[j].src==bin){ found=true; break; }

          if(!found && covers.indexOf(bins[i].id.value)==-1) unused+=" � "+bins[i].id.value+"\n";
 }

 if(unused==''){ if(bins && bins.length>0) unused=' � ��� ������������\n'; else unused=' � ��� �������� ��������\n'; }

 unused="\n�������������� �������� �������: \n"+unused;

 // Find broken images (no binary)

   if(!imgs || imgs.length==0, !bins || bins.length==0){ MsgBox("\n� ���� ����� ��� �� �������, �� �����������. \n\n"); return; }

 var broken="";

 for(var i=0; i<imgs_len; i++)
 {
  var found=false; var img=imgs[i].src.substr(14);

  for(var j=0; j<bins_len; j++) if(bins[j].id.value==img){ found=true; break; } 

  if(!found) broken+=" � "+img+"\n";
 }

 if(broken=='') broken=" � �� �������\n";

 alert("           �= AlexS Script=� \n   ��������� ����������� v.1.1\n\n" +unused+"\n������������� �����������: \n"+broken+"\n");

}

