# url2ppt
Generate a PPT file by a URL ,using Deepseek LLM.

The Python script is supposed to :

(1)Crawl the content of some web page.

(2)Analysis main themes and key points of the article by **Deepseek** LLM.

(2)Generate a PPT file by **python-pptx**.

Before running the url2ppt.py, pls set  **DEEP_API_KEY** in the enviroment variables , 

```
set DEEP_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

and install essential dependencies first.

```
pip install json requests bs4 python-pptx PIL pytesseract urllib newspaper sklearn backoff
```
 
**Usage:**

```
python url2ppt.py [URL]
```
