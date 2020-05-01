from docx import Document
import re


document = Document('protocol53knesset22.docx')


def readDocument(document):
    isSpeech = False
    isChair = False
    speakerName = ""

    chairWords = 0
    otherWords = 0
    rows = []

    for p in document.paragraphs:
        if p.text.strip() == "": continue
        x = re.search("<<([^>]+)>>(.*)<<[^>]+>>\s*", p.text); # title regex
        if x != None: # if it is a title
            if x[1] == u" יור ":
                speakerName = x[2]
                isSpeech = True
                isChair = True
            elif x[1] == u" דובר " or x[1] == u" דובר_המשך " or x[1] == u" קריאה " or x[1] == u" אורח ":
                speakerName = x[2]
                isSpeech = True
                isChair = False
            else:
                isSpeech = False
        elif isSpeech:
            row = {"isChair": isChair , "chairWords": 0, "otherWords":0, "speakerName": speakerName, "text":p.text.strip()}
            if isChair:
                row["chairWords"] = len(row["text"].split(" "))
                chairWords += row["chairWords"]
                print(p.text)
            else:
                row["otherWords"] = len(row["text"].split(" "))
                otherWords += row["otherWords"]
            rows.append(row)

    results = {"chairWords": chairWords, "otherWords":otherWords, "total": chairWords + otherWords}
    results["chairWordsPercentage"] = chairWords * 100 / results["total"]
    results["otherWordsPercentage"] = otherWords * 100 / results["total"]
    return (rows, results)
    #print(chairWords * 100 / (otherWords + chairWords))
