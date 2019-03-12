# Python
最後編輯日 2018/11/26

# VS Code
F2           <--- 重新命名
F12          <--- 移至定義
ALT+SHIFT+F  <--- 自動排版
ALT+0        <--- 全部摺疊
ALT+SHIFT+0  <--- 全部展開

# Git
git init                                 <--- 目錄初始化
git config --global user.name            <--- 查詢 Git Hub 帳號
git config --global user.email           <--- 查詢 Git Hub 帳號 Email
git clone https://github.com/Tragicalone <--- 複製 Tragicalone 的 Git Hub 到目前資料夾
git add 檔案名                           <--- 將檔案加入 Git 追蹤
git commit -m "訊息"                     <--- 同意檔案的變更並加註訊息 
git push                                 <--- 將同意的變更上傳至 Git Hub
git log --oneline --graph                <--- 版本歷史流程圖
git diff
git reflog
git reset --hard HEAD^^
git reset --hard TheNumber
git checkout TheBranch
git merge --no-ff -m "Accept message" TheBranch
git remote -v

#Jupyter
pip install "ipython[notebook]"   <--- 安裝 Jupyter
pip install opencv-python         <--- 安裝 OpenCV
jupyter notebook                  <--- 於目前目錄啟動 Jupyter Service
