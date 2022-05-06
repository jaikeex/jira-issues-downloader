import os
import requests
from datetime import datetime
from requests.auth import HTTPBasicAuth
import pandas as pd
import warnings

request_params = {
    "url": "https://jira.ixperta.com",
    "username": "",
    "password": "",
    "issue_keys": [],
    "ignore_internal_comments": True
}


invalidating_text_snippets = [
    "_THIS IS AN INTERNAL IXPERTA COMMENT FOR PURPOSE OF SLA NOTIFICATION._"
]


data = {
    "Issue": [],
    "Text_Type": [],
    "Text": [],
    "Date": [],
    "Author": []
}


dataframe = pd.DataFrame(data)


def main():
    print("Starting...")
    warnings.simplefilter(action='ignore', category=FutureWarning)
    filepath = os.path.join(str(os.getcwd()), "output.xlsx")

    dfm = get_final_dataframe()
    dfm.to_excel(filepath)

    print(f"Success! File can be found at {filepath}")


def get_final_dataframe():
    # customers = {"Remote JIRA Sync User (J2J)"}  -- for fleetcor issues
    customers = {""} 

    get_user_input()

    for issue_key in request_params["issue_keys"]:
        print(f"Processing Issue key={issue_key}.")

        issue_data = get_issue_data_from_jira(issue_key)

        customers.add(issue_data["customer"])
        comments = get_comments(issue_data["comments"], issue_data["customer"])

        dataframe = add_summary(issue_data, dataframe)
        dataframe = add_description(issue_data, dataframe)
        dataframe = add_comments(issue_data, dataframe, comments)

    dataframe = apply_styling(customers=customers, issue_keys=request_params["issue_keys"])
    return dataframe


def get_issue_data_from_jira(issue_key):
    issue_url = request_params["url"] + "/browse/" + issue_key
    url = request_params["url"] + "/rest/api/2/issue/" + issue_key

    issue = get_request(url)
    issue_data = parse_issue(issue_key, issue_url, issue)
    return issue_data


def parse_issue(issue_key, issue_url, issue):
    issue_data = {
            "excel_link": f'=HYPERLINK("{issue_url}", "{issue_key}")',
            "date": datetime.strptime(issue["fields"]["created"][:19], "%Y-%m-%dT%H:%M:%S"),
            "summary": issue["fields"]["summary"],
            "description": issue["fields"]["description"],
            "customer": issue["fields"]["reporter"]["displayName"],
            "comments": issue["fields"]["comment"]["comments"]
        }
    
    return issue_data


def apply_styling(customers, issue_keys):
    print("Applying styling to the dataframe.")
    dataframe = dataframe.reset_index()
    dataframe = dataframe.style.apply(
        df_summary_background_style,
        target="summary",
        column=["Text_Type"],
        axis=1).applymap(df_bold_style, customers=customers)
    dataframe = dataframe.applymap(df_link_style, subset=["Issue"])
    dataframe = dataframe.applymap(df_link_style_un, subset=["Issue"])

    return dataframe


def add_comments(issue_data, comments):
    for comment in comments:
        comment_row = {
            "Issue":     issue_data["excel_link"],
            "Text_Type": "comment",
            "Text":      comment["text"],
            "Date":      comment["date"],
            "Author":    comment["author"]
        }

        dataframe = dataframe.append(comment_row, ignore_index=True)
    return dataframe


def add_description(issue_data):
    description_row = {
        "Issue":     issue_data["excel_link"],
        "Text_Type": "description",
        "Text":      issue_data["description"],
        "Date":      issue_data["date"],
        "Author":    issue_data["customer"]
    }
    dataframe = dataframe.append(description_row, ignore_index=True)
    return dataframe


def add_summary(issue_data):
    summary_row = {
        "Issue":     issue_data["excel_link"],
        "Text_Type": "summary",
        "Text":      issue_data["summary"],
        "Date":      issue_data["date"],
        "Author":    issue_data["customer"]
    }
    dataframe = dataframe.append(summary_row, ignore_index=True)
    return dataframe


def get_comments(all_comments, customer):
    comments = []
    for comment in all_comments:
        comment_data = parse_comment(comment)
        print(f"Processing comment id={comment_data['id']}")

        ignored_comment = is_ignored_comment(comment_data["text"], comment_data["id"])

        if not ignored_comment:    
            format_comment_body(comment_data)
            add_comment_to_list_of_all_comments(customer, comments, comment_data)

    return comments


def format_comment_body(comment_data):
    if comment_data["text"].startswith("_commented by "):
        comment_data["text"] = ''.join(comment_data["text"].splitlines(keepends=True)[1:])


def parse_comment(comment):
    comment_data = {
            "text": comment["body"],
            "id": comment["id"],
            "date": datetime.strptime(comment["created"][:19], "%Y-%m-%dT%H:%M:%S"),
            "author": comment["author"]["displayName"]
        }
    
    return comment_data


def add_comment_to_list_of_all_comments(customer, comments, comment_data):
    if comment_data["author"] == customer:
        comments.append({
            "text":        comment_data["text"],
            "date":        comment_data["date"],
            "author":      comment_data["author"],
            "author_type": "customer",
        })
    else:
        comments.append({
            "text":        comment_data["text"],
            "date":        comment_data["date"],
            "author":      comment_data["author"],
            "author_type": "agent",
        })


def is_ignored_comment(comment_text, comment_id):
    url = request_params["url"] + "/rest/api/2/comment/" + comment_id + "/properties/sd.public.comment"
    visibility = get_request(url)

    ignore_comment = False

    if "value" in visibility and request_params["ignore_internal_comments"]:
        if visibility["value"]["internal"] == True: 
            ignore_comment = True
            print(f"Comment id={comment_id} is marked as internal. It will be skipped.")

    for invalidating_text in invalidating_text_snippets:
        if invalidating_text in comment_text:
            ignore_comment = True
            print(f"Comment id={comment_id} contains {invalidating_text}, marking it as inavalid. It will be skipped.")

    return ignore_comment


def get_request(url):
    basic = HTTPBasicAuth(request_params["username"], request_params["password"])
    response = requests.get(url, auth=basic)
    return response.json()


def get_user_input():
    url_input = input("Enter JIRA base url. Defaults to https://jira.ixperta.com if left blank: ")
    if len(url_input) > 0:
        request_params["url"] = url_input

    username_input = input("Enter your username: ")
    request_params["username"] = username_input

    password_input = input("Enter your password: ")
    request_params["password"] = password_input

    issue_keys_input = input("Enter issue keys separated by spaces: ")
    request_params["issue_keys"] = issue_keys_input.split(" ")


def df_link_style(val):
    return "color: blue"


def df_link_style_un(val):
    return "text-decoration: underline"


def df_bold_style(val, customers):
    return "font-weight: bold" if str(val) in customers else ""


def df_summary_background_style(val, target, column):
    is_summary = pd.Series(data=False, index=val.index)
    is_summary[column] = val.loc[column] == target
    return ['background-color: orange' if is_summary.any() else '' for v in is_summary]


if __name__ == '__main__':
    main()