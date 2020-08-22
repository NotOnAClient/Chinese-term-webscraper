import bs4, requests, docx, chinese
text = docx.Document("file.docx") #read file.docx
try:
    output = docx.Document("output.docx") #read output.docx if exists, else creates a new one
except:
    output = docx.Document()
clear = input("Clear output.docx?(y/n): ")
analyzer = chinese.ChineseAnalyzer() #required for looking up chinese characters in strings

def main(text):
    url = "https://baike.baidu.com/item/"
    print(text.text)
    res = requests.get(url + text.text, headers={'User-Agent': 'Mozilla/5.0'}) #must use headers, or some websites will not work
    print(res.url)
    res.encoding = "utf-8" #makes cjk readable on python IDLE console
    res.raise_for_status()
    baiduSoup = bs4.BeautifulSoup(res.text, "lxml") #lxml is some sort of html parser
    define = baiduSoup.find("meta", {"name":"description"}) #get meta, name=description
    definition = define["content"] if define else None #get attribute 'content' from meta
    #print(define["content"] if define else None)
    if definition != None: #takes one sentence from the entire paragraph
        sen = analyzer.parse(definition)
        if sen.search("即"):
            x = sen.search("即")
            print(x)
            return x
        if sen.search("指"):
            x = sen.search("指")
            print(x)
            return x
        else:
            x = sen.sentences()
            print(x[0])
            return x
    else:
        pass

def clear_document(doc): #clears the entire word document, no idea how it works and never will, copied off the internet somewhere
    for a in doc.paragraphs:
        p = a._element
        p.getparent().remove(p)
        p._p = p._element = None
        
while True: #ask for clearing document       
    if clear == 'y':
        clear_document(output)
        break
    elif clear == 'n':
        break
    else:
        clear = input("Clear output.docx?(y/n): ")

for i in text.paragraphs:
    y = main(i)
    try:
        output.add_paragraph(i.text + " - " + str(y[0]) + "。")
    except TypeError:
        output.add_paragraph(i.text + " - " + "Definition not found on Baidu")
    print("-------------------------------------------------------------------------")
output.save("output.docx")
input("Done!")
