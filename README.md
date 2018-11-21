# OutputConfig
Output excel config files to json or binary files

## Features
- Support three common data structures **array/dictionary/struct** and primitive types **int8/int16/int32/bool/float/string**
- Easy definition, a dictionary: *dict\<int, string\>*, an array: *array\<int\>*, define a struct in **struct.json** and reference it with *&* like *&custom_struct*
- Support any combination of data structure and primitive type, such as *array\<dict\<array\<&custom_struct\>,dictionary\<&custom_struct, string\>\>\>*

## Data in Excel


!!cs|c|cs|cs
---|---|---|---
int|string|dict\<string,dict\<&int_array_wrapper,array\<dict\<int,string\>\>\>\>|dict\<&test_strcut,array\<int\>\>
ID desc	|This is name| |data1 desc
ID|name|data|data1	
1|test|<ad:<{[0,0,1]}:[<1:adfadf,2:adfwe>,<3:adsf>] >>|<{<[asdf,asdfa]:1,[dsfa]:123>,<<1231:123,453:adfa>:[{1,45,2},{1241}],<1231:123,453:adfa>:[{1,45,2},{1241}]>,<[asdf,rqwtq,",<>fasd[]{}"]:124>}:[124]>
2|34531|<asdf:<{[0,0,1]}:[<1:adfadf,2:adfwe>,<3:adsf>] >>|<{<[asdf,asdfa]:1,[dsfa]:123>,<<1231:123,453:adfa>:[{1,45,2},{1241}],<1231:123,453:adfa>:[{1,45,2},{1241}]>,<[asdf,rqwtq,",<\">fasd[]{}"]:124>}:[124,124,523]>

- The first row represents the type of data column, **c** means client, **s** means server. One excel can export different type of configs with differrent flags.
- **!!** means the start point of data, so the data can be in anywhere of the excel just only need starting with !! in the row.
- The second row represents the type definition of data column.
- The third row represents the comment of data field, or just left it blank.
- The fourth row represents the name of data field.
- Every config file has an ID field, as the primary key of all data.

## How to use it
- Define a custom struct in struct.json, *class_name* means the C# class name, *define* means all the fields of the struct, *is_generate* measn whether or not to generate the C# class file of the struct
```
"chapter_reward":{
    "class_name": "ChapterReward",
    "define": { 
      "star1": "array<&int_array_wrapper>",
      "star2": "dict<&int_array_wrapper, string>",
      "star3": "array<&int_array_wrapper>"
    },
    "is_generate": 1 
  }
```
- Run the test.xlsx file in python: `python OutputConfig.py test.xlsx`. An Unity project directory will be automatically generated, open the Main.unity in Unity and enjoy it!
- Both the .json and .bytes will be generated. Json format doesn't support a dictionary with non string key and Unity's native json decode functions don't support multidimensional array, so I suppose you to use the .bytes file most of the time.

