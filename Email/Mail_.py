import smtplib 
ob=smtplib.SMTP("smtp.gmail.com",587)
ob.starttls()
ob.login("himanshumanral2003@gmail.com","qvicfgwfprwbnlrd")
subject="hii there for you .."
body="This is a system genreted email...."
mes="subject:{}\n\n{}".format(subject,body)
print(mes)
listofadd=["deepmanral2004@gmail.com","nareshsinghankit@gmail.com"]
ob.sendmail("himanshumanral2003@gmail.com",listofadd,mes)
print("send successfully !")
ob.quit()