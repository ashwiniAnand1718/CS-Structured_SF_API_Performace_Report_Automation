import requests

class download_PerformanceResult():
    """Please provide a directory path there the performance raw result file needs to be downloaded &
     it should contain 2 folders Windows & Linux
     Ex:"C:\\Users\\sa\\Documents\\PythonCode\\Downloaded_report\\Windows"""


download_directory ="C:\\Users\\sa\\PycharmProjects\\PT_report_formatting\\Download_result\\"
 # Please provide the jenkins login credentails
username = "qins"
password = "Rdis2fun"
base_url = "http://oaf-lxsavqa305:8080/view/WSAAPI/job/"
end_url = "/ws/scripts/results/"
jenkins_jobName = ["SFAPI-Performance-Linux-Regression-Static","SFAPI-Performance-Linux-Regression-CreditModel","SFAPI-Performance-Windows-Regression-Static","SFAPI-Performance-Windows-Regression-CreditModel"]
file_name= ["analysis_results_all_summary_static.xlsx","analysis_results_all_summary_credit_model.xlsx"]
os_version = ["Windows", "Linux"]
run_model = ["Static", "Credit"]


def download_reults_file(full_directory,full_url):
    response = requests.get(full_url, auth=(username, password))
    with open(full_directory, "wb") as f:
        f.write(response.content)

#def rename_files():


for x in os_version:
    for y in jenkins_jobName:
        if y.__contains__(x) and y.__contains__(run_model[0]):
            final_url = base_url+y+end_url+str(file_name[0])
            full_directory_path=download_directory+x+"\\"+str(file_name[0])
            print(x,y)
            download_reults_file(full_directory_path,final_url)

        elif y.__contains__(x) and y.__contains__(run_model[1]):
            final_url = base_url + y + end_url+str(file_name[1])
            full_directory_path=download_directory+x+"\\"+str(file_name[1])
            print(x,y)
            download_reults_file(full_directory_path,final_url)











































"""for i in file_name:
    for x in jenkins_jobName:
        for j in os_version:
            for s in run_model:
                if j==("Linux") and x.__contains__("Linux"):
                   if x.__contains__(s):
                        final_url = base_url + x + end_url + i
                        directory_name=download_directory_Linux+str(i)
                        download_reults_file(directory_name, final_url)
                elif j==("Windows") and x.__contains__("Windows"):
                    if x.__contains__(s):
                        final_url = base_url + x + end_url + i
                        directory_name = download_directory_Windows + str(i)
                        download_reults_file(directory_name, final_url)


print("All The file are downloaded successfully")"""






""""  for j in os_version:
            if x=="Linux" : 
                final_url= base_url+j+end_url+file_name
                download_reults_file(download_directory_Linux,final_url)
elif x== "Windows" :
final_url = base_url + j + end_url + file_name
download_reults_file(download_directory_Windows,final_url)"""


