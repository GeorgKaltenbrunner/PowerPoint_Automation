# PowerPoint_Automation

Den Link fand ich sehr hilfreich: https://towardsdatascience.com/creating-presentations-with-python-3f5737824f61


## Main.py

- Ich find es am hilfreichsten mich in das pptx Package einzulesen
- Zeile 8 - 15: Hier wird die Titelfolie erstellt
- Danach (in Kommentar gepackt) kommt dein Inhalt rein. Den Inhalt kannst du dann über prs.slide_layouts[index] reinpacken. Der Index bezieht sich auf den Folienmaster und die Foliennummer dort
- Wenn du Bilder reinpacken magst, fand ich add_image.py ganz hilfreich

## add_image.py

- Letztendlich wird hier nur geschaut, dass das Bild nicht größer als die Folie ist und einigermaßen einheitlich aussieht
- placeholder_id ist theoretisch auch wieder aus dem Folienmaster rauslesbar. Weiß gerade aber nicht mehr, wo ich die gefunden habe :/
