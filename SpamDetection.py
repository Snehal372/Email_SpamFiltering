import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
from win32com.client import Dispatch

def speak(text):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(text)


model = pickle.load(open('spam.pkl','rb'))
cv=pickle.load(open('vectorizer.pkl','rb'))


def main():
    st.title("Email Spam Classification Application")
    st.subheader("Build with Streamlit and Python")
    message=st.text_input("Enter a Text: ")
    if st.button("Process"):
        data= [message]
        vect= cv.transform(data).toarray()
        prediction= model.predict(vect)
        result= prediction[0]
        if result==1 :
            st.error("This is a Spam Email")
            speak("This is a Spam Email")
        else :
            st.success(" This is a Ham Email")
            speak("This is a Ham Email")
main()