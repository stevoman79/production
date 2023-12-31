from docx import Document
import requests
from datetime import date, datetime
from apscheduler.schedulers.blocking import BlockingScheduler
#from apscheduler.schedulers.background import BlockingScheduler


def extract_program_data(file_path):
    doc = Document(file_path)
    program_data = []

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells

            # Check if the first cell contains the date information
            if len(cells) >= 3 and cells[0].text.strip():
                date = cells[0].text.strip()
                time = cells[1].text.strip()
                event = cells[2].text.strip()

                # Format date and time as 'yyyy-mm-dd hh:mm:ss'
                formatted_datetime = f"{date[6:]}-{date[3:5]}-{date[:2]} {time}"
                program_data.append([formatted_datetime, event])

    return program_data

# Provide the path to your Word document
file_path = "bushman.docx"

# Extract program data from the Word document
program_data = extract_program_data(file_path)

# Access the extracted data
formatted_datetimes = [data[0] for data in program_data]
events = [data[1] for data in program_data]

# Print the formatted datetimes
"""
print("Formatted Datetimes:")
for formatted_datetime in formatted_datetimes:
    print(formatted_datetime)

# Print the events
print("Events:")
for event in events:
    print(event)
"""
print("\n")
print(program_data[1:])
print("\n")
for i in program_data[1:]:
    print(i)


##################################################

def sendWSP(message, apikey, gid=0):
    url = "https://whin2.p.rapidapi.com/send"
    headers = {
        "content-type": "application/json",
        "X-RapidAPI-Key": apikey,
        "X-RapidAPI-Host": "whin2.p.rapidapi.com"
    }
    try:
        if gid == 0:
            return requests.request("POST", url, json=message, headers=headers)
        else:
            url = "https://whin2.p.rapidapi.com/send2group"
            querystring = {"gid": gid}
            return requests.request("POST", url, json=message, headers=headers, params=querystring)
    except requests.ConnectionError:
        return "Error: Connection Error"


def send_message(msg, api_key):
    my_message = {"text": msg}
    sendWSP(my_message, api_key)


# Schedule the messages using APScheduler
def schedule_messages():
    sched = BlockingScheduler()

    # Define the dates and messages to be sent
    schedule_data = program_data[1:]
        
        #('2023-12-27 18:31:00', "Happy New Year!"),
        #('2023-12-27 18:32:00', "Happy Valentine's Day!"),
        #('2023-12-27 18:33:00', "Happy St. Patrick's Day!")
    #]

    # Schedule each message
    for date, message in schedule_data:
        sched.add_job(send_message, 'date', run_date=date, args=[message, myapikey])

    sched.start()


if __name__ == "__main__":
    myapikey = "ae31ae9ebfmsh5862ad5122f501dp18bd02jsn040cc6b2a095"
    schedule_messages()

