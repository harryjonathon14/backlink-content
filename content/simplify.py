"""Light copy-editing: keep technical meaning and grammatical number intact."""
import re
WORDS={
    'additionally':'also','additional':'extra',
    'requirements':'rules','requirement':'rule',
    'proposed':'planned','proposal':'bid','proposals':'bids',
    'maintenance':'care',
    'regularly':'often','frequently':'often',
    'understandable':'clear',
    'approximately':'about','approximate':'rough',
    'verify':'check','verified':'checked',
    'obtain':'get','receive':'get','received':'got',
    'primarily':'mainly','numerous':'many',
    'immediately':'right away',
    'regarding':'about',
    'excavation':'digging','excavating':'digging','excavated':'dug-up','excavate':'dig',
    'completed':'finished',
    'practical':'useful',
    'potential':'possible','persistent':'ongoing',
    'restrict':'limit','restricted':'limited','restrictions':'limits',
    'underlying':'root','equivalent':'equal',
    'initial':'first','ordinary':'normal',
    'significant':'major','substantial':'major',
}
PATTERN=re.compile(r'\b('+ '|'.join(sorted(map(re.escape,WORDS),key=len,reverse=True))+r')\b',re.I)
def simplify(text):
    def sub(m):
        value=WORDS[m.group().lower()]
        return value[0].upper()+value[1:] if m.group()[0].isupper() else value
    text=PATTERN.sub(sub,text)
    # Repair articles only for words introduced by the safe substitutions above.
    text=re.sub(r'\b([Aa])n (first|normal|major|rough|useful|finished|planned|bid|rule|root)\b',r'\1 \2',text)
    text=re.sub(r'\b([Aa]) (extra|equal)\b',lambda m:('An' if m.group(1)=='A' else 'an')+' '+m.group(2),text)
    text=re.sub(r', but ','. But ',text)
    text=re.sub(r', so ','. So ',text)
    text=re.sub(r', and (you|it|they|we)\b',lambda m:'. '+m.group(1).capitalize(),text)
    return text
