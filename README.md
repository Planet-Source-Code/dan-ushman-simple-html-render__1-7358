<div align="center">

## Simple HTML Render


</div>

### Description

Shows HTML... For example, if you use it on "< b>blah< /b>" it would prinkt BLAH in bold letters.

So far only supports bold, italic and underline html tags... I dont wnat to work on this anymore, i was bored when i made this. heh
 
### More Info
 
Picturebox, html code

Rendered Html onto the picturebox.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Ushman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-ushman.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-ushman-simple-html-render__1-7358/archive/master.zip)





### Source Code

```
'
' A simple demo on how to write a HTML renderer.
' written by Dan Ushman <ushman@mediaone.net>
' Please visit Refsoft at www.refsoft.com
'
' This code is free. It is not restricted in ANY way
' you can use it, take credit for it, do what ever you want
' with it. I honestly don't care.
'
' I know its not perfect, but I did not spend to much
' time on it. I wrote this in 10 minutes one night
' when I was bored.
'
' Anyway, please E-mail me and tell me what you think
' And...
'
' Enjoy,
' Dan Ushman - ushman@mediaone.net - www.refsoft.com
'
Option Explicit     'Many programmers do not use this. What they dont know is
            'weather or not they declare there variables can and will
            'have a large effect on how much memory your program will use
            'and how stable it will be. I recommend that every one
            'use this line of code, and declare every variable they use
            'I learned this the hard way, while writting Uut I was wondering why
            'it took so much ram... Well , thats all.
Sub RenderHTML(pic As PictureBox, html As String)
  '
  ' Always declare variables
  '
    'Integers
    Dim lentext As Integer
    Dim html_loop_1 As Integer 'The main loop
    Dim html_loop_2 As Integer 'Secondary loop
    Dim html_pos_1 As Integer  'Opening carret
    Dim html_pos_2 As Integer  'Closing carret
    'Strings
    Dim str_html As String   'The copy of the original HTML string
    Dim html_tag As String   'Stores the tag...
    Dim html_text As String   'Stores the text to be modified by the tags
    Dim cur_char As String   'Used in the loops, one char at a time
    'Boolean
    Dim open_c As Boolean    'Is it an opening carret?
    Dim close_c As Boolean   'Is it a closing carret?
  '
  ' Get the length of the HTML and some other things...
  '
    lentext = Len(html)     'The length of the HTML string
    str_html = html       'The copy of the original HTML string
  '
  ' Loop though the HTML
  '
    For html_loop_1 = 1 To lentext         'The main loop
      html_pos_1 = InStr(str_html, "<")      'Find the locations of the Opening and Closing carrets
      html_pos_2 = InStr(str_html, ">")
      cur_char = Mid(str_html, html_loop_1, 1)  'Go though the HTML byte by byte
      If cur_char = "<" Then           'Is it an openning carret?
        open_c = True
        close_c = False
        html_tag = ""              'Clear the tag variable, for now.
      ElseIf cur_char = ">" Then         'Maby not...
        open_c = False
        close_c = True
        If InStr(html_tag, "<") Then
          html_tag = Right(html_tag, Len(html_tag) - InStr(html_tag, "<"))
        End If
      End If
      If open_c = True And close_c = False Then    'If the carret is currently open...
        html_tag = html_tag & cur_char       'combine all the chrs after it until the carret closes...
      End If                     'I am very sure there are tons of better ways to do this,
                              'but this works fine.
      If close_c = True And open_c = False Then
        If Not cur_char = "<" And Not cur_char = ">" Then
          html_text = html_text & cur_char    'Add each char together aslong as its not a carret (both kinds) or
        End If                   'part of a tag. This part could use some work, its not perfect and is rather buggy.
      End If
      '
      'So far this little project of mine only supports BOLD, ITALIC and UNDERLINE HTML tags. I may or may not
      'add more support. I am lazy, so don't bet your dinner.
      '
      If close_c = True And open_c = False Then
        html_tag = LCase(html_tag)         'Make sure the tag is lowercase.
        Select Case html_tag            'Start going though the tag, and doing what it wants us to do
          Case Is = "b"
            pic.FontBold = True         'If the tag is on, make the text bold, else dont...
          Case Is = "i"
            pic.FontItalic = True
          Case Is = "u"
            pic.FontUnderline = True
          Case Is = "/b"
            pic.FontBold = False
          Case Is = "/i"
            pic.FontItalic = False
          Case Is = "/u"
            pic.FontUnderline = False
        End Select
        pic.Print html_text;
        html_text = ""               'Clear the variables when we are done.
        html_tag = ""
      End If
    Next html_loop_1                  'And we are on our way... again.
End Sub
```

