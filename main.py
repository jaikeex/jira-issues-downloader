from os import link
import requests
from datetime import datetime
from requests.auth import HTTPBasicAuth
import pandas as pd

request_params = {
    "url": "https://jira.ixperta.com",
    "username": "",
    "password": "",
    "issue_keys": []
}

data = {
    "Issue": [],
    "Text_Type": [],
    "Text": [],
    "Date": [],
    "Author": []
}

def get_data_from_jira(dataframe):
    customers = {""}
    issue_keys = {""}

    get_user_input()

    basic = HTTPBasicAuth(request_params["username"], request_params["password"])
    for issue_key in request_params["issue_keys"]:
        issue_url = request_params["url"] + "/browse/" + issue_key
        link = f'=HYPERLINK("{issue_url}", "{issue_key}")'

        issue_keys.add(issue_key)
        issue = get_issue_json(basic, issue_key)

        date = datetime.strptime(issue["fields"]["created"][:19], "%Y-%m-%dT%H:%M:%S")
        summary = issue["fields"]["summary"]
        description = issue["fields"]["description"]
        all_comments = issue["fields"]["comment"]["comments"]

        customer = issue["fields"]["reporter"]["displayName"]
        customers.add(customer)

        comments = get_comments(all_comments, customer)

        dataframe = add_summary(customer, dataframe, date, link, summary)
        dataframe = add_description(customer, dataframe, date, description, link)
        dataframe = add_comments(comments, dataframe, link)

    dataframe = apply_styling(dataframe, customers=customers, issue_keys=issue_keys)

    return dataframe


def apply_styling(dataframe, customers, issue_keys):
    dataframe = dataframe.reset_index()
    dataframe = dataframe.style.apply(
        df_summary_background_style,
        target="summary",
        column=["Text_Type"],
        axis=1).applymap(df_bold_style, customers=customers)
    dataframe = dataframe.applymap(df_link_style, subset=["Issue"])
    dataframe = dataframe.applymap(df_link_style_un, subset=["Issue"])

    return dataframe


def add_comments(comments, dataframe, link):
    for comment in comments:
        comment_row = {
            "Issue":     link,
            "Text_Type": "comment",
            "Text":      comment["text"],
            "Date":      comment["date"],
            "Author":    comment["author"]
        }

        dataframe = dataframe.append(comment_row, ignore_index=True)
    return dataframe


def add_description(customer, dataframe, date, description, issue_key):
    description_row = {
        "Issue":     issue_key,
        "Text_Type": "description",
        "Text":      description,
        "Date":      date,
        "Author":    customer
    }
    dataframe = dataframe.append(description_row, ignore_index=True)
    return dataframe


def add_summary(customer, dataframe, date, issue_key, summary):
    summary_row = {
        "Issue":     issue_key,
        "Text_Type": "summary",
        "Text":      summary,
        "Date":      date,
        "Author":    customer
    }
    dataframe = dataframe.append(summary_row, ignore_index=True)
    return dataframe


def get_comments(all_comments, customer):
    comments = []
    for comment in all_comments:
        comment_author = comment["author"]["displayName"]
        comment_date = datetime.strptime(comment["created"][:19], "%Y-%m-%dT%H:%M:%S")
        if comment_author == customer:
            comments.append({
                "text":        comment["body"],
                "date":        comment_date,
                "author":      comment_author,
                "author_type": "customer",
            })
        else:
            comments.append({
                "text":        comment["body"],
                "date":        comment_date,
                "author":      comment_author,
                "author_type": "agent",
            })
    return comments


def get_issue_json(basic, issue_key):
    url = request_params["url"] + "/rest/api/2/issue/" + issue_key
    response = requests.get(url, auth=basic)
    issue = response.json()
    return issue


def main():
    df = pd.DataFrame(data)
    dfm = get_data_from_jira(df)
    dfm.to_excel("d:/output.xlsx")


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