# Outlook Replied Emails to Training Data

Reads N number of your sent email conversations, converts them to training data, then writes to a .txt file.

* <a href="#info">General Information</a>
* <a href="#training">Training COLAB

## <a id="info" href="#toc">General Information</a>
* Email reply bodies are assumed to follow the default format:
```
From: NAME <EMAIL>\r\nSent: DATE\r\nTo: USER\r\nSubject: SUBJECT\r\n \r\nBODY
```
* Currently formats the data into the following format:
```
{
    "prompt": [
        [
            {
                "role": "user",
                "content": "From: 'SENDER NAME <SENDER EMAIL>' To: 'USER NAME <USER EMAIL>' Sent Date: 'DATE' With subject: 'SUBJECT' With content: 'BODY'"
            }
        ]
    ],
    "completion": [
        [
            {
                "role": "assistant",
                "content": "BODY"
            }
        ]
    ]
}
```

## <a id="training" href="#toc">Training COLAB</a>
* Data can be trained using this GITHUB with Unsloth:
  - https://colab.research.google.com/drive/1fKWyiiDJoWy_hGqETYHfQiYiRX5cVKx6?usp=sharing


