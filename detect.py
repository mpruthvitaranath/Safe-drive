from win32com.client import Dispatch
import cv2
import dlib
from scipy.spatial import distance
import imutils
from imutils import face_utils


#Function to Calculate Eye Aspect Ratio

def ratio(eye):
    A = distance.euclidean(eye[1] , eye[5])
    B = distance.euclidean(eye[2] , eye[4])
    C = distance.euclidean(eye[0] , eye[3])
    EAR = (A + B )/(2.0*C)
    return EAR

# setting thresholds

threshold = 0.25
time_thresh = 25

detector = dlib.get_frontal_face_detector()
shape_predict = dlib.shape_predictor(r"C:\Users\taranath\Desktop\shape_predictor_68_face_landmarks.dat")

(lStart,lEnd) = face_utils.FACIAL_LANDMARKS_IDXS["left_eye"]
(rStart,rEnd) = face_utils.FACIAL_LANDMARKS_IDXS["right_eye"]


#voice of PC 

speak = Dispatch("SAPI.Spvoice")

#Initializing the web cam

cam = cv2.VideoCapture(0)

rand = 0
while True:
    ret , frame = cam.read()
    frame = imutils.resize(frame , width = 500)
    gray = cv2.cvtColor(frame , cv2.COLOR_BGR2GRAY)
    faces = detector(gray,0)
    for face in faces:
        shape = shape_predict(gray,face)
        shape = face_utils.shape_to_np(shape)
        lefteye = shape[lStart:lEnd]
        righteye = shape[rStart:rEnd]

        #Calculating the Eye aspect ratio for each eye and averaging it

        leftear = ratio(lefteye)
        rightear = ratio(righteye)
        ear = (leftear + rightear )/ 2.0

        #convex hull

        lefthull = cv2.convexHull(lefteye)
        righthull = cv2.convexHull(righteye)

        #drawing the contours

        cv2.drawContours(frame , [lefthull] , -1 , (0,255,0) , 2)
        cv2.drawContours(frame , [righthull] , -1 , (0,255,0) , 2)

        if ear < threshold:
            rand += 1
            print (rand)
            if rand >= time_thresh:
                cv2.putText(frame, "!!!!!!!!ALERT!!!!!!!", (200, 30),cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)

                #adding PC voice
                speak.Speak("Be alert Warning
                            ")
        else:
            rand = 0

    cv2.imshow('Detector',frame)
    if cv2.waitKey(1) == 13:
        break

cv2.destroyAllWindows()
