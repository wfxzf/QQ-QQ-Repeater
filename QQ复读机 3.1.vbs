' Copyright (C) 2020  wfxzf

' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.


Option Explicit

dim wshshell,count,span,paste,counter,i

set wshshell=createobject("wscript.shell")

count = inputbox ("����Ҫ���͵Ĵ���"&chr(13)&"��ȷ���ղŹ������QQ������!","QQ������3.0")
span = inputbox("����ÿ�η��͵�ʱ��������λ�����루ms��","QQ������3.0")
paste = msgbox("�Ƿ�ճ�����а����ݣ�",vbyesno,"QQ������3.0")

if paste=vbyes then
    counter = msgbox("�Ƿ���м�����",vbyesno,"QQ������3.0")
end if

WshShell.AppActivate ""

wscript.sleep 1000 

for i=1 to count 
    wscript.sleep span 
    if paste=vbyes then
        wshshell.sendkeys "^v"

        if counter = vbyes then
            wshshell.sendkeys "+9"
            wshshell.sendkeys i
            wshshell.sendkeys "+0"
        end if 
    else 
        wshshell.sendkeys i
    end if
    wshshell.sendkeys "%s"
next
wscript.quit