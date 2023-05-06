import cv2
import imutils
import extract
import pytesseract

pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

def binarization(image):
    thresh, im_bw = cv2.threshold(image, 178, 243, cv2.THRESH_BINARY)
    return im_bw

img_file = "sinhvien/maihuyhoang2.jpg"
image = cv2.imread(img_file)
image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
image = binarization(image)
data = pytesseract.image_to_string(image, lang="vie")
print(data)
boxes = pytesseract.image_to_data(image)
print(boxes)
for x, b in enumerate(boxes.splitlines()):
    if x != 0:
        b = b.split()
        if len(b) == 12:
            l, t, w, h = int(b[6]), int(b[7]), int(b[8]), int(b[9])
            cv2.rectangle(image, (l, t), (l + w, h + t), (0, 0, 255), 2)
cv2.imshow("Scanned", imutils.resize(image, height=650))
cv2.waitKey(0)
cv2.destroyAllWindows()
name = ''
student_id = ''
date = ''
course = ''
department = ''
major = ''
score_sm1 = ''                  # score semester I
score_sm2 = ''                  # score semester II
score_avg = ''                  # avg score per semester

lines = data.splitlines()

dict_scr_sm1 = ["học kì I: ", "học kì I;", "học kì J;", "học kì I ", "học kì L ", "học kì I-", "học kì : "]
dict_scr_sm2 = ["học kì II: ", "học kì II ", "học kì JI ", "học kì II; ", "học kì L :", "học kì JI: ", "học kì JL ", "học kì 1L ", "học kì lI: "]

for line in lines:
    if score_sm1 == '':
        for i in dict_scr_sm1:
            if i in line:
                score_sm1 = extract.find_score(line, i)
    if score_sm2 == '':
        for j in dict_scr_sm2:
            if j in line:
                score_sm2 = extract.find_score(line, j)
    if score_avg == '':
        score_avg = extract.find_score(line, "học:")
    if name == '':
        name = extract.find_name(line)
    if department == '':
        department = extract.find_department(line)
    if major == '':
        major = extract.find_major(line)

student_id = extract.find_studentID(data)
course = extract.find_course(data)
date = extract.find_date(data)

#kết quả trích xuất ảnh
print("Name: ", name)
print("Student ID: ", student_id)
print("Date of Birth: ", date)
print("Course: ", course)
print("Department: ", department)
print("Major: ", major)
print("Score semester I: ", score_sm1)
print("Score semester II: ", score_sm2)
print("Average score: ", score_avg)