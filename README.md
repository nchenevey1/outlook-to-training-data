# Outlook Replied Emails to Training Data
Reads N number of your sent email conversations, converts them to training data, then writes to a .txt file.

# VBA Query Model and Query Model Script
Reads the currently selected email and queries a local model using llama.cpp to create a reply email draft.

* <a href="#info">General Information</a>
* <a href="#training">Training COLAB</a>
* <a href="#demo">Demonstration</a>

## <a id="info" href="#toc">General Information</a>
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
  - https://colab.research.google.com/drive/1pJLuRyHyV4TnMTtYcL-tbxA65fWACcMK?usp=sharing

## <a id="demo" href="#toc">Demonstration</a>

https://github.com/user-attachments/assets/203ad6f3-12f4-4214-b504-1035c999c608