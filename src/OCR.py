import tabula
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt




def extract_df(file_path):
    df = tabula.read_pdf(file_path, pages='all')
    size = len(df)
    data = df[0]
    for i in range(1,size-1):
        data = pd.concat([data, df[i]])
    data["Reference"] = data["Reference"].astype(str)
    data["Reference"] = [x if ('PM' in x) else None for x in data["Reference"] ]
    data = data.dropna()
    data["Date"]=data["Date"].astype('datetime64[ns]')
    data["Claimed Amount"]=[x.replace(',','') for x in data["Claimed Amount"]]
    data["Claimed Amount"]=data["Claimed Amount"].astype(float)
    data["Total Aprroved"]=[x.replace(',','') for x in data["Total Aprroved"]]
    data["Total Aprroved"]=data["Total Aprroved"].astype(float)
    data["Weekday"] = [x.weekday() for x in data["Date"]]
    data["Approval"] = [True if x==y else False for x,y in zip(data["Claimed Amount"], data["Total Aprroved"])]
    return data

def plot_summary(df):
    fix, axs = plt.subplots(2,2)
    sns.countplot(df["*Status"], ax=axs[0,0])
    axs[0,0].set(xlabel='Payment Status')
    sns.countplot(df["Weekday"], ax=axs[0,1])
    axs[0,1].set(xlabel='Days of the Week')
    axs[0,1].set_xticks(ticks=range(7),labels=['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'])
    sns.histplot(df["Claimed Amount"], ax=axs[1,0])
    axs[1,0].set(xlabel='Payment Amount')
    sns.countplot(df["Approval"], ax=axs[1,1])
    axs[1,1].set(xlabel='Approval Rate')
    plt.show()

if __name__ == "__main__":
    file_path = ".\\lib\\test.pdf"  
    df = extract_df(file_path)
    plot_summary(df) 
    input("Press Enter to end...")

