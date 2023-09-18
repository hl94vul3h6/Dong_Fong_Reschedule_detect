# Dong_Fong_Reschedule_detect<br>

## Alert: Delivery Date Change Notification<br>
Detect orders from the daily Excel data exported from the ERP where the delivery date (scheduled delivery date) has been modified, and there's a potential for late delivery. Finally, email the results to all sales representatives, allowing them to verify orders they are responsible for based on their sales representative ID.<br>

#### Attachment Description:<br>
1. **email.txt**: Used to modify the sender and recipient emails. See instructions within the file for input methods.<br>
2. **fileName.txt**: To modify the names of the two files you wish to compare. Refer to the instructions inside the file for input methods.<br>

#### Data Processing Explanation:<br>
1. Assume the ERP exports a month of order data daily, and the expected completion date = expected off-machine date + 10 days.<br>
2. After reading two consecutive days of data, use the 'pk' field to identify duplicate data and handle it, identifying any changes made to the expected delivery or off-machine dates.<br>
3. Extract data with changed expected delivery dates.<br>
4. Identify data that might be late (where the expected completion date exceeds the agreed delivery date).<br>
5. Consolidate the information needed by sales representatives, convert it to HTML, and send via email.<br>

#### Program Logic Explanation:<br>
1. pk (unique) = Order Number + Item Number<br>
2. df = Yesterday's data [4/1-5/1] = [4/1 + 4/2-5/1 (Before change + Unchanged)]<br>
3. df1 = Today's data [4/2-5/2] = [5/2 + 4/2-5/1 (After change + Unchanged)]<br>
4. tmp [4/1 + 5/2 + 4/2-5/1] = Yesterday's data + Today's data (All new rows added)<br>
5. all_diff [4/1 + 5/2 + 4/2-5/1] = All unique data between yesterday and today (including orders with the same 'pk' before and after change)<br>
6. pk_diff [4/1 + 5/2 + 4/2-5/1 (Unchanged)] = All data between yesterday and today with unique 'pk' values (excluding all modified orders)<br>
7. first [4/1 + 4/2-5/1 (Before change + Unchanged)] = All data with unique 'pk' values (including pre-modified orders, excluding post-modified orders)<br>
8. last [5/2 + 4/2-5/1 (After change + Unchanged)] = All data with unique 'pk' values (including post-modified orders, excluding pre-modified orders)<br>
9. old [4/2-5/1 (Before change)] = All modified data from yesterday<br>
10. new [4/2-5/1 (After change)] = All modified data from today<br>
11. merged_data [4/2-5/1 (Before + After change)] = All data from two days where delivery date or off-machine date was modified, displaying both days' dates separately.<br>
12. changed_data [4/2-5/1 (Changed delivery date)] = All data with modified delivery dates over two days.<br>
13. delayed_data [4/2-5/1 (Changed delivery date & potential delay)] = Data from 'changed_data' where the expected completion date > agreed delivery date.<br>
14. final = The final notification content to be emailed to sales reps, selected from 'delayed_data' with columns reformatted as needed.<br>

#### Order Test File Explanation:<br>
1. Day 1 data: **Order test.xlsx**<br>
2. Day 2 data:<br>
   1. **Order test_No Notification.xlsx** (Shouldn't receive an email)<br>
   2. **Order test_Notify.xlsx** (Should receive notifications for Items 1 and 3)<br>

#### Each item's data representation:

| Item | Scheduled Delivery Date | Expected Off-Machine Date | Potential Late Delivery | Send Notification |
|:----:|:-----------------------:|:-------------------------:|:-----------------------:|:-----------------:|
|  1   |       Changed           |       Changed             |         Yes             |        Yes        |
|  2   |       Changed           |       Changed             |         No              |        No         |
|  3   |       Changed           |       Unchanged           |         Yes             |        Yes        |
|  4   |       Changed           |       Unchanged           |         No              |        No         |
|  5   |       Unchanged         |       Changed             |         Yes             |        No         |
|  6   |       Unchanged         |       Changed             |         No              |        No         |
|  7   |       Unchanged         |       Unchanged           |         Yes             |        No         |
|  8   |       Unchanged         |       Unchanged           |         No              |        No         |



# Dong_Fong_Reschedule_detect_中文版解釋

## Alert: 交期更改通知
找出ERP每天轉出的Excel訂單資料中，交期 (排定交貨日) 有修改且該訂單可能遲交的的數據。再將最後結果 email 給所有業務，業務可由業務編號自行核對是否有自己負責的訂單。 <br>

#### 附件說明
1.email.txt: 用來修改寄件人與收件人email，輸入方式請看檔案內文。<br>
2.fileName.txt: 用來修改欲比較的兩個檔案名稱，輸入方式請看檔案內文。<br>

#### 資料處理流程說明
1. 假設ERP每天轉出一個月的訂單資料、且預計完工日 = 預計下機日 + 10天。<br>
2. 讀取連續兩天的data後，利用欄位pk判斷資料重複並進行處理，可得出所有預計交貨日或下機日被更改過的資料。<br>
3. 取預計交貨日有更改的data出來。<br>
4. 取有可能遲交(預計完工日超過約定交貨日)的data。<br>
5. 將業務需要的資訊整合並轉html，再以email寄出。<br>

#### 程式邏輯說明
1.	pk (不能重複) = 訂單編號 + 項次<br>
      步驟2到11中的更改是指交期或下機日被更改
2.	df = 昨天的資料 [4/1-5/1] = [4/1 + 4/2-5/1(更改前+未改的)]
3.	df1 = 今天的資料 [4/2-5/2] = [5/2 + 4/2-5/1(更改後+未改的)]
4.	tmp [4/1 + 5/2 + 4/2-5/1(更改前+更改後+未改的*2)]<br>
      = 昨天的資料 + 今天的資料 (全部以新row加入)<br>
      = [4/1 + 4/2-5/1(更改前+未改的)] + [5/2 + 4/2-5/1(更改後+未改的)]<br>
      = df + df1<br>
5.	all_diff [4/1 + 5/2 + 4/2-5/1(更改前+更改後+未改的)]<br>
      = 昨天跟今天所有不重複的資料 (含更改前後pk相同的訂單)<br>
      = 昨天的資料 + 今天的資料 - 兩天中每個column都一樣的資料<br>
      = [4/1 + 5/2 + 4/2-5/1(更改前+更改後+未改的*2)] – [4/2-5/1未改的]<br>
      = tmp - tmp中每個column都一樣的資料
6.	pk_diff [4/1 + 5/2 + 4/2-5/1(未改的)]<br>
      = 昨天跟今天所有pk不重複的資料 (不含被更改過的所有訂單)<br>
      = all_diff - all_diff中pk一樣時的所有資料<br>
      = all_diff刪除所有pk一樣的資料<br>
7.	first [4/1 + 4/2-5/1(更改前+未改的)]<br>
	= 昨天跟今天所有pk不重複的資料 (含更改前的訂單，不含更改後的訂單)<br>
      = all_diff - all_diff中pk一樣時今天的資料<br>
      = all_diff中pk一樣時保留昨天的資料，刪除今天的<br>
8.	last [5/2 + 4/2-5/1(更改後+未改的)]<br>
	= 昨天跟今天所有pk不重複的資料 (含更改後的訂單，不含更改前的訂單)<br>
      = all_diff - all_diff中pk一樣時昨天的資料<br>
      = all_diff中pk一樣時保留今天的資料，刪除昨天的
9.	old [4/2-5/1更改前的]<br>
      = 昨天所有被更改過的資料<br>
       =[4/1 + 4/2-5/1(更改前+未改的)] + [4/1 + 4/2-5/1未改的] – [5/2 + 4/2-5/1未改的]<br>
       = first + pk_diff – 兩者pk一樣的資料<br>
10.	new [4/2-5/1更改後的]<br>
      = 今天所有被更改過的資料<br>
       = [5/2 + 4/2-5/1(更改後+未改的)] + [5/2 + 4/2-5/1未改的] – [5/2 + 4/2-5/1未改的]<br>
      = last + pk_diff – 兩者pk一樣的資料<br>
11.	merged_data[4/2-5/1更改前+後的]<br>
      = 兩天中所有被更改過交期或下機日的資料，並將兩天的交期與下機日分別顯示<br>
12.   changed_data[4/2-5/1交期被更改的]<br>
       = 兩天中所有被更改過交期的資料
13.	delayed_data[4/2-5/1交期被更改且可能遲交的]<br>
      = changed_data中，預計完工日 > 約定交貨日的資料<br>
      = 交期有被更改且可能來不及交貨的資料<br>
14.	final = 最後要寄給業務的通知內容<br>
      = 從delayed_data中選擇業務需要的內容並將欄位格式重新調整<br>

#### 訂單test檔案說明
1. 第一天的資料：訂單test.xlsx
2. 第二天的資料：
    1. 訂單test_不通知.xlsx (不應收到信)
    2. 訂單test_要通知.xlsx (應收到項次1與3的通知)<br>
    3. 各項次數據代表意義

| 項次 | 排定交貨日 | 預計下機日 | 可能遲交 | 發通知 |
|:----:|:----------:|:----------:|:--------:|:------:|
|   1  |    改變    |    改變    |    是    |   是   |
|   2  |    改變    |    改變    |    否    |   否   |
|   3  |    改變    |    不變    |    是    |   是   |
|   4  |    改變    |    不變    |    否    |   否   |
|   5  |    不變    |    改變    |    是    |   否   |
|   6  |    不變    |    改變    |    否    |   否   |
|   7  |    不變    |    不變    |    是    |   否   |
|   8  |    不變    |    不變    |    否    |   否   |
