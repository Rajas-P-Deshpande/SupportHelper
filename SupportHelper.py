import time
import csv
import nltk
from tkinter import *
#import Tkinter as tk
from nltk.tokenize import sent_tokenize, word_tokenize
#nltk.download('punkt')
import docx2txt
import operator
import re, math
from collections import Counter
from nltk.corpus import stopwords
from nltk import pos_tag
from nltk.corpus import wordnet as wn
import os


def get_cosine(vec1, vec2):
         intersection = set(vec1.keys()) & set(vec2.keys())
         numerator = sum([vec1[x] * vec2[x] for x in intersection])

         sum1 = sum([vec1[x]**2 for x in vec1.keys()])
         sum2 = sum([vec2[x]**2 for x in vec2.keys()])
         denominator = math.sqrt(sum1) * math.sqrt(sum2)

         if not denominator:
            return 0.0
         else:
            return float(numerator) / denominator

    
def text_to_vector(text):
         words = WORD.findall(text)
         return Counter(words)


def penn_to_wn(tag):
    #print("penn_to_wn")
    """ Convert between a Penn Treebank tag to a simplified Wordnet tag """
    if tag.startswith('N'):
        return 'n'
 
    if tag.startswith('V'):
        return 'v'
 
    if tag.startswith('J'):
        return 'a'
 
    if tag.startswith('R'):
        return 'r'
 
    return None
def tagged_to_synset(word, tag):
    #print("tagged_to_synset")
    wn_tag = penn_to_wn(tag)
    if wn_tag is None:
        return None
 
    try:
        return wn.synsets(word, wn_tag)[0]
    except:
        return None
def sentence_similarity(sentence1, sentence2):
    """ compute the sentence similarity using Wordnet """
    # Tokenize and tag
    #print("sentence_similarity")
    sentence1 = pos_tag(word_tokenize(sentence1))
    sentence2 = pos_tag(word_tokenize(sentence2))
 
    # Get the synsets for the tagged words
    synsets1 = [tagged_to_synset(*tagged_word) for tagged_word in sentence1]
    synsets2 = [tagged_to_synset(*tagged_word) for tagged_word in sentence2]
 
    # Filter out the Nones
    synsets1 = [ss for ss in synsets1 if ss]
    synsets2 = [ss for ss in synsets2 if ss]
 
    score, count = 0.0, 0
 
    # For each word in the first sentence
    for synset in synsets1:
        # Get the similarity value of the most similar word in the other sentence
        
        try:
            best_score = max([synset.path_similarity(ss) for ss in synsets2])
        except:
            best_score = 0
 
        # Check that the similarity could have been computed
        if best_score is not None:
            score += best_score
            count += 1
 
    # Average the values
    #if count :
    try:
        score /= count
    except:
        score = 0.0
    #else:
        #score = 0
    return score

'''with open('Sampleinput.csv', 'r') as csvfile:
    spamreader = csv.reader(csvfile)
    #for row in spamreader:
        #print(', '.join(row))
    line2line = {}
    lines=[]
    for line in spamreader:
        lines.append(line)
    
    for ele in lines:
            line2line[ele[4]]=ele[5]
    '''
WORD = re.compile(r'\w+')
line2line = {}
data = dict()
with open('dict.txt') as raw_data:
    for item in raw_data:
        if ':' in item:
            key,value = item.split(':', 1)
            line2line[key]=value

#print(line2line)
#print (line2line['Account locked. Not able to login'])


try:
    data = input("Enter Your Query:")
except:
    print("Error!")
    
data1 = data
avg_similarity, c  = 0.0 , 0
osg = {}
for key,text2 in line2line.items():
        data= ' '.join([word for word in data.lower().split() if word not in stopwords.words("english")])
        #print(data)
        vector1 = text_to_vector(data)
        originalkey = key
        key= ' '.join([word for word in key.lower().split() if word not in stopwords.words("english")])
        vector2 = text_to_vector(key)

        cosine = get_cosine(vector1, vector2)
        #if(cosine>0):
          #  print("cosine value",cosine)
           # print(key)
            
        if data1 == key:
            c += 1
            osg[c]=text2
            print("option:"+str(c))
            print(originalkey,':',text2) 
            print()
            break
        else:
            value = sentence_similarity(data1, key)
            if  value > 0.2 or cosine > 0.2 :
                c += 1
                print("option:"+str(c))
                osg[c]=text2
                print(originalkey,':',text2)
                #print('simi val:',value)
                #print('avg:',value+(cosine/2.5))
                print()

if data1!=key:
    print("option:"+str(c+1))
    
    token_data=word_tokenize(data1)   
    text = docx2txt.process("sampleiod1.docx")
    text2 = docx2txt.process("sampleiod2.docx")
    text3 = docx2txt.process("sampleiod3.docx")
    f = open('file.txt','w')
    f.write(text)
    f.close()
    f = open('file.txt','a+')
    f.write("\r\n")
    f.write(text2)
    f.write("\r\n")
    f.write(text3)
    f.close()
    filename='file.txt'
    file = open(filename, 'rt')
    text = file.read()
    file.close()
    sentence=sent_tokenize(text)
    #stop_words_sent = set(stopwords.sentence('english'))
    #sentence = [w for w in sentence if not w in stop_words_sent]
    #print(sentence)
    key=0;
    my_Dict = {}
    for line in sentence:
        count=0
        for x in token_data:
            if x.lower() in line.lower():
                count=count+1
        my_Dict[key]=count
        key=key+1
    data_index=max(my_Dict.items(), key=operator.itemgetter(1))[0]
    print("Your query: "+str(sentence[data_index]))
    
    #print(text)
#print(c)
helpfuls = {}
try:
    option=input("Choose your option:")
except:
    print("Error!")
#print(c)
if((int(option))==c+1):
    print("For more information refer the file just got opned.......")
    os.startfile('file.txt')
else:
    try:
        response=input("Was it helful...yes/no:")
    except:
        print("Error!")
    if(str(response)=="yes" or str(response)=="y"):
        print("you are welcome!!")
        helpfuls[data1]=osg[int(option)]

#print(line2line)
with open('helpfuls.txt', 'w') as f:
    for key, value in helpfuls.items():
        f.write('%s:%s\n' % (key, value))
