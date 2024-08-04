from generate_outline_template import LLOutlineTemplate,MLOutlineTemplate,HLOutlineTemplate
from generate_research_template import LLResearchTemplate,MLResearchTemplate,HLResearchTemplate
import os
import shutil
import filecmp

list_number=0
force_regenerate=True

'''EDIT SECTION'''

def generate_templates(question_file):
    total_pictures=10
    total_videos=5

    topic=question_file
    topic=topic.replace(".txt","")
    topic=topic.replace("_"," ")
    topic=topic.title()

    presentation_file_name="LL Presentation Template.pptx"

    test_version=question_file
    test_version=test_version.replace("_"," ")
    if not topic.upper() in test_version.upper():
        print(topic,test_version)
        print("WARNING: NAME AND TOPIC DO NOT MATCH")
        confirmation=input("Are you sure you want to proceed? Enter y for yes. ")
        if "y" in confirmation:
            print("Proceeding")
        else:
            return

    '''END OF EDIT SECTION'''
    questions=[]
    with open("questions/"+question_file,"r") as f:
        lines=f.readlines()
        for i in range(len(lines)):
            line=lines[i]
            line=line.strip()
            line=line.replace("\n","")
            questions.append(line)
    
    start_folder="Python_Generated"
    topic=topic.replace("-"," ")
    topic=topic.replace("_"," ")
    topic_plusgenerated=f"({start_folder}) {topic} Organized"
    levels=["LL","ML","HL"]
    for level in levels:
        folder_name=f"{start_folder}/{topic_plusgenerated}/{topic} {level}"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

            #Copy presentation file to each folder
            shutil.copy(presentation_file_name,f"{folder_name}/3 - {presentation_file_name}")

    folder_name=f"{start_folder}/{topic_plusgenerated}/{topic} LL"
    research_ll=LLResearchTemplate("LL",topic,total_pictures,total_videos,questions,folder_name)
    research_ll.generate_research_documents()

    folder_name=f"{start_folder}/{topic_plusgenerated}/{topic} ML"
    research_ml=MLResearchTemplate("ML",topic,total_pictures,total_videos,questions,folder_name)
    research_ml.generate_research_documents()

    folder_name=f"{start_folder}/{topic_plusgenerated}/{topic} HL"
    research_hl=HLResearchTemplate("HL",topic,total_pictures,total_videos,questions,folder_name)
    research_hl.generate_research_documents()

    folder_name=f"{start_folder}/{topic_plusgenerated}/{topic} LL"
    outline_ll=LLOutlineTemplate("LL",topic,total_pictures,total_videos,questions,folder_name)
    outline_ll.generate_outline_documents()

    folder_name=f"{start_folder}/{topic_plusgenerated}/{topic} ML"
    outline_ml=MLOutlineTemplate("ML",topic,total_pictures,total_videos,questions,folder_name)
    outline_ml.generate_outline_documents()

    folder_name=f"{start_folder}/{topic_plusgenerated}/{topic} HL"
    outline_hl=HLOutlineTemplate("HL",topic,total_pictures,total_videos,questions,folder_name)
    outline_hl.generate_outline_documents()

def main():
    #If the file_names are too long, there are issues with finding the files
    for question_file in os.listdir("questions"):
        start_location=f"questions/{question_file}"
        end_location=f"updated_questions/{question_file}"
        generating=False

        #If there is not an end file, update the soft skills research
        if not os.path.exists(end_location):
            generating=True
        #If the original and end file do not match, update the soft skills research
        elif not filecmp.cmp(start_location,end_location,shallow=False):
            generating=True

        if force_regenerate:
            generating=True

        if generating:
            print(question_file)
            generate_templates(question_file)
            shutil.copy(start_location,end_location)


if __name__=="__main__":
    main()