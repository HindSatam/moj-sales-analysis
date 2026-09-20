

1- What is in here: coverage by district, by period, by category
The finding shows that the data has 4 Distinct cities (جدة, دوقه, كبشان, الرياض)
With 126 null values.


2- What is wrong with the data, and why?


Null Vlaues:
The statistic summary shows inconsistency in the (non-null count) which indicate null values in the data.
The column with the larger number of null values of 806 is "RS_TYPE".
The smaller number of null values of 120 is in the column "SERIAL".
It appears that a value of arabic null is exisiting in the "NEIGHBORHOOD_NAME" column; 'لايوجد'.


Duplicates:
There is 14215 row with identical SERIAL number, which considered a problem in the data (assuming every SERIAL number identify one uniqye record). But since it is big part of the data set, we will keep them assuming it is a normal behavior.
This finding led me to investigate more into the duplicates value against SERIAL column. 
It shows that there is SERIAL values appear in multiple rows with distinct content (in other column): 6117 record
And rows that are identical in everything except SERIAL number: 2625 record.


Text hygiene:
Some records in the "RS_TYPE" column has values with square bracket, and some is number.
"new_use" also has number as values.
I noticed that some values in "NEIGHBORHOOD_NAME" convey the same meaning with different Arabic syntax such as: 
[الاجواد, الأجواد]
[الامير فواز الجنوبى, الأمير فواز الجنوبى]
[الربوه, الربوة]
[الصوارى, الصواري]

Outliers;
There is outliers on the extremes that affect the data and needs to be solved to avoid incorrect values.


3- Which of those problems would change a published average sale price or total sales value, and by roughly how much

It seems like the oulier make it harder to read into the data and compare different problems, after dropping the possible outliers from the extremes, we got a more reasonable results.
Extreme values change the raw average by about 100%. Other then that, duplicates inflate total value by about 2%, and the other problems are negligible.