# MS Word files indexer
[Created in August 2015]

I recently co-edited a book for a major publisher. As academic books go, there’s usually an index at the end of it to help readers getting to the key information quickly. Indexing for books is actually quite an art form, and publishers used to hire indexers in swathes. Nowadays, alas, they are only hired by publishers if the book promises to be a best-seller — which our book probably ain’t. 😦

Which meant we had to compile the index ourselves — or pay a freelance indexer to do the job for us. Unfortunately, this was not within our budget. So I thought about ways to come up with a little software app to help me in this endeavor. Given that the time frame was pretty short, I didn’t have much time to delve into this too deeply. I was thus happy to have an imperfect script, that would give me something to work with.

As I have done some NLP in the past, I figured some code to extract the nouns and verbs from all chapters in the book should suffice to have an initial list which I could then edit manually to get a list of index terms. This page (http://www.howtogeek.com/howto/35495/how-to-create-an-index-table-like-a-pro-with-microsoft-word/) explains how to generate an index with a "concordance document" (i.e. a document with a table of index terms) in Word. Given that I had every chapter of the book in a separate file, the outline of my approach was thus:

    1) Read all *.docx files in a subdirectory
    2) For each of those, extract the key terms to go into the index
    3) Collate all these terms into a single concordance list
    4) Write out this list as docx file

Now, I have had some experience with NLTK to do NLP. However, to get good results with it requires a bit of sophistication. As mentioned, I was happy to get imperfect results. While google’ing for solutions by other folks, I stumbled upon Topia.termextract (https://pypi.python.org/pypi/topia.termextract/), which admittedly produces a lot of noise, but is very easy to use. So, the idea was to use the python-docx library (https://python-docx.readthedocs.org/en/latest/) to read all the chapters, extract the key terms with Topia.termextract, and then write out the concordance file with python-docx again. 

There’s certainly room for improvement here. To begin with, one could experiment with Topia’s filters. I have commented them out in the code, as they limited the list too much. The result without filter contains a lot of garbage, but I needed something real quick, and going through the results manually was quicker than trying to find a better algorithm.

Maybe if I got time at some point, I’ll come up with something more sophisticated. Using NLTK with some heuristics to get only terms that make sense for a book index might be an idea, maybe by checking resources such as Wordnet or Dbpedia to infer meaningful key words…
