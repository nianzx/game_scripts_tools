python实现游戏后台脚本的库

原本是打算用python来实现，写好工具后准备要写脚本时，试了下打包，发现几个工具类打包出来都要100M+，放弃了。改用其他语言。。。

- **README.md**  
- **requirements.txt**  		python依赖库
  
    - **dll**  
        - leptonica-1.82.0.dll  	ocr用
        - tesseract50.dll  			ocr用
  
    - **tessdata**  
        - chi_sim.traineddata  		ocr的语言包
  
    - **utils**  
        - GraphColorUtil.py  		图色命令
        - KeymouseUtil.py    		键鼠命令
        - OCR.py  			 		ocr命令
        - TimeUtil.py   	 		延时用
        - WindowsUtil.py	 		窗口命令
		
写出来的脚本支持后台运行，只要你的游戏窗口不是最小化就可以 你可以边挂机边看视频