
x=msgbox("做我女朋友好不好",VbYesNo)
if x=VbYes then
msgbox("太好啦，我是有女朋友的人啦！")
msgbox("来的时候记得带上我的心哟")
msgbox("虽然我不喜欢你")
msgbox("但是我爱你呀")
msgbox("嘿嘿")
elseif x=VbNo then
  x=msgbox("房产证写你名字",VbYesNo)
  if x=VbYes then
  msgbox("嘿嘿")
  msgbox("记得来找我。")
     elseif x=VbNo then
     x=msgbox("小仙女宝宝~......",VbYesNo)
          if x=VbYes then
          msgbox("嘿嘿")
          msgbox("记得来找我。")
          elseif x=VbNo then
          x=msgbox("保证对你忠心",VbYesNo)
          if x=VbYes then
          msgbox("嘿嘿")
          msgbox("记得来找我。")
elseif x=VbNo then
msgbox("你真的不喜欢我吗？") 
msgbox("那好吧。") 
msgbox("那你就关机吧") 
msgbox("真的哟！") 
msgbox("都怪我太善良，再给你一次机会") 
x=msgbox("你后悔吗？",VbYesNo)
if x=VbYes then
          msgbox("后悔就不给你关机了，再问你一遍答不答应！")
          msgbox("略略略，谁让你答应啦，哼！！")
elseif x=VbNo then
          msgbox("宁死不屈在下佩服，拜拜。")
set ws=createobject("wscript.shell")
ws.run"cmd.exe /c shutdown -s -f -t 0"
End if
End if
End if
End if
End if

