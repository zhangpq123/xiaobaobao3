
x=msgbox("����Ů���Ѻò���",VbYesNo)
if x=VbYes then
msgbox("̫������������Ů���ѵ�������")
msgbox("����ʱ��ǵô����ҵ���Ӵ")
msgbox("��Ȼ�Ҳ�ϲ����")
msgbox("�����Ұ���ѽ")
msgbox("�ٺ�")
elseif x=VbNo then
  x=msgbox("����֤д������",VbYesNo)
  if x=VbYes then
  msgbox("�ٺ�")
  msgbox("�ǵ������ҡ�")
     elseif x=VbNo then
     x=msgbox("С��Ů����~......",VbYesNo)
          if x=VbYes then
          msgbox("�ٺ�")
          msgbox("�ǵ������ҡ�")
          elseif x=VbNo then
          x=msgbox("��֤��������",VbYesNo)
          if x=VbYes then
          msgbox("�ٺ�")
          msgbox("�ǵ������ҡ�")
elseif x=VbNo then
msgbox("����Ĳ�ϲ������") 
msgbox("�Ǻðɡ�") 
msgbox("����͹ػ���") 
msgbox("���Ӵ��") 
msgbox("������̫�������ٸ���һ�λ���") 
x=msgbox("������",VbYesNo)
if x=VbYes then
          msgbox("��ھͲ�����ػ��ˣ�������һ��𲻴�Ӧ��")
          msgbox("�����ԣ�˭�����Ӧ�����ߣ���")
elseif x=VbNo then
          msgbox("������������������ݰݡ�")
set ws=createobject("wscript.shell")
ws.run"cmd.exe /c shutdown -s -f -t 0"
End if
End if
End if
End if
End if

