# PDF大量更名指令行工具
1. 本工具已編譯版本為Windows版本
2. 本工具設計目的在於讀入一個規則檔，依據內容或是PDF檔案的中介（metadata）更名到指定的檔案，例如某個PDF檔案中身分證號會出現2次，那就更名為該身分證號
3. 請務必使用本工具附帶的csv檔案來更名，如果要自己建立，欄位定義請參考附帶的檔案，然後你自己要記得把csv存成UTF8格式（Excel 2016之前似乎都沒有這個能力）
4. 本工具有兩個步驟，第一步為把一個大檔案分割，第二步為更名，如果你沒有要做第一步，按Enter可以直接跳過
5. 這只是個簡單的小程式，採MIT授權，歡迎轉載，改寫記得標註我即可
6. 其實本來打算用golang寫，可惜golang沒有好用的pdf套件，於是又回來寫C#，很多年沒碰了有什麼邏輯不順請見諒