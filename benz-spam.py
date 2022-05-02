import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
import win32com.client
# from win32com.client import Dispatch

hide_menu = """
<style>
#MainMenu{
  visibility:hidden}

footer{
	visibility:hidden
}  
footer:after {
	content:'Copyrights © BENZ-Email Spam Classification App'; 
	visibility: visible;
	display: block;
	position: relative;
	#background-color: red;
	padding: 5px;
	top: 2px;
}
</style>
"""

def speak(text):
	# speak=Dispatch(("SAPI.SpVoice"))
	speak = win32com.client.Dispatch("SAPI.SpVoice")
	speak.Speak(text)


model = pickle.load(open('spam.pkl','rb'))
cv=pickle.load(open('vectorizer.pkl','rb'))


def main():
	st.title("Email Spam Classification App")
	st.markdown(hide_menu,unsafe_allow_html=True)
	
	st.write("Build with Python NLP MultinomialNB, Streamlit")
	activites=["main","about"]
	choices=st.sidebar.selectbox("Select Activities",activites)
	if choices=="main":
		st.subheader("From: [Prajwal Benedict A](https://www.linkedin.com/in/prajwal-benedict-a-048511186/)")
		# st.subheader("Classification")
		msg=st.text_input("Enter a text")
		if st.button("Process"):
			print(msg)
			print(type(msg))
			data=[msg]
			print(data)
			vec=cv.transform(data).toarray()
			result=model.predict(vec)
			if result[0]==0:
				st.success("This is Not A Spam Email")
				speak("This is Not A Spam Email")
			else:
				st.error("This is A Spam Email")
				speak("This is A Spam Email")

	if choices=="about":
		st.subheader("Made with ♥ by Prajwal benedict A")
		st.write("check out the [git repo](https://github.com/BENEDICT9845/)")
main()