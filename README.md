# Converting images into spreadsheets
This program uses openpyxl, numpy and opencv to convert an image into cells in a spreadsheet.

### How?
- Each row of pixels in the image is represented in the spreadsheet by a blue green and red row
- Each cell in the spreadsheet then gets a value based on its color
- Each row is formatted so that the value of the cells defines the color of the cells
- ![Zoomed in view](https://i.ibb.co/GQRKmfz/IMAGES-IN-EXCEL.png)
- Zoom out far enough and you start seeing the full picture
- ![Full view](https://i.ibb.co/NCvCpqp/image.png)

### Credit
- This is not an idea i came up with
- I saw this from Matt Parker on his youtube channel [standupmaths](https://www.youtube.com/user/standupmaths)