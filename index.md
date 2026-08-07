+++
draft = false
+++

 

<style>
/*   表格存在/Users/max/project/watermelonMax/static/data  */
input{

    color:#222;

}
/* 内容区域 */

.word-box{

    background:rgba(255,255,255,0.9);
    
    width:1400px;
    
    padding:40px;
    
    margin:60px auto;
    
    border-radius:15px;
    
    box-shadow:0 5px 20px rgba(0,0,0,0.08);
    
    text-align:center;

    position:relative;

}


/* 按钮 */

button{

    padding:10px 30px;
    
    font-size:20px;
    
    cursor:pointer;
    
    border:none;
    
    border-radius:3px;
    
    background:#4a90e2;
    
    color:white;

    margin-top:5px;

}


button:hover{

    background:#357abd;

}


/* 单词 */

h1{

    color:#222;

    margin-top:80px;

}
label{
    font-size:22px;
    color:#222;
    gap:15px;
    margin-top:60px;
    
}
#word-list{

    display:grid;
    
    grid-template-columns:repeat(5,1fr);
    
    gap:15px;
    
    margin-top:60px;
    
    font-size:22px;
    
    color:#111827;

}



.word-card{

    background:rgba(255,255,255,0.75);
    
    border-radius:10px;
    
    padding:20px 10px;
    
    min-height:70px;
    
    display:flex;

    flex-direction:column;

    align-items:center;

    justify-content:center;

    text-align:center;

}

body{

    font-family:Arial,sans-serif;
    
    background-image:url('/images/1.webp');
    
    background-size:cover;
    
    background-position:center;
    
    background-attachment:fixed;

}

#word-count-info{

    font-size:18px;
    
    color:#555;
    
    margin-bottom:20px;

}
 
.seed-input{

    width:80px;

 

}

.answer-input{

    width:80%;

    padding:8px;

    font-size:18px;

    border-radius:6px;

    border:1px solid #aaa;

}
.result{

    margin-top:15px;

    color:#2563eb;

    font-size:20px;

    font-weight:bold;

}
.speak-btn{

    background:transparent;

    border:none;

    color:#2563eb;

    font-size:22px;

    padding:2px 5px;

}
.pomodoro{

    position:absolute;

    top:30px;

    right:30px;


    width:260px;

    padding:20px;


    /* 半透明玻璃 */

    background:rgba(255,255,255,0.35);


    /* 毛玻璃效果 */

    backdrop-filter:blur(15px);


    border-radius:20px;


    /* 悬浮阴影 */

    box-shadow:0 20px 50px rgba(0,0,0,0.3);


    z-index:1000;


    text-align:center;


    /* 浮动动画 */

    animation:float 3s ease-in-out infinite;
    
}
#timer{


    font-size:70px;

    font-weight:bold;


    color:white;


    text-shadow:

    0 3px 10px rgba(0,0,0,0.4);


}


.progress{


    height:6px;

    background:

    rgba(9, 8, 8, 0.3);


    border-radius:5px;


    margin:15px 0;

    rgba(0, 0, 0, 0.3);

    transform-origin:left;


}
.progress{
    display:flex;
    justify-content:flex-end;
}

#progress-bar{


    height:100%;


    width:100%;


    background:#ffffff;


    border-radius:5px;

    rgba(148, 19, 19, 0);


}

.timer-buttons button{


    padding:8px 12px;


    font-size:18px;


    background:

    rgba(0, 0, 0, 0.3);


}
@keyframes float{


    0%{

        transform:translateY(0px);

    }


    50%{

        transform:translateY(-4px);

    }


    100%{

        transform:translateY(0px);

    }

}

.robot{

    position:absolute;

    right:40px;

    bottom:20px;

}


.robot lottie-player{

    width:150px;

    height:150px;

}

#top-btn{

    position:fixed;

    right:40px;

    bottom:90px;

    width:50px;

    height:50px;

    background:transparent;
 
 

}


#top-btn span{

    display:block;

    width:42px;

    height:42px;

    border-left:5px solid white;

    border-top:5px solid white;

    transform:rotate(45deg);

    margin:15px auto 0;

    filter:drop-shadow(0 2px 3px black);

}

 

</style>
 


 
<script src="https://unpkg.com/@lottiefiles/lottie-player@latest/dist/lottie-player.js"></script> 
<script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>

 
<div class="word-box">
 
 

<div class="pomodoro">

 
<div id="timer">
25:00
</div>


<div class="progress">
    <div id="progress-bar"></div>
</div>


<div class="timer-buttons">

<button onclick="startTimer()">
▶
</button>


<button onclick="pauseTimer()">
Ⅱ
</button>


<button onclick="resetTimer()">
↻
</button>



</div>


</div>

<h1>
英语大比拼
</h1>


<div id="word-count-info">
词库数量：加载中...
</div>

<label>
随机种子：
</label>
<input 
id="seed"
type="number"
class="seed-input"
value="20">

<label>
单词数量：
</label>
<input 
id="word-count"
type="number"
value="20"
min="1"
max="100"> 

 
<button onclick="test()">
随机起来！
</button>

 

<div id="word-list"></div>


{{< robot >}}

  

<script>



let words = [];
let wrongWords = [];


//保存当前随机出来的单词
let currentWords = [];
//随机种子
let seed = null;
 
//读取普通词库

fetch('/data/words.xlsx?t=' + Date.now())

.then(response=>response.arrayBuffer())

.then(data=>{


    let workbook=XLSX.read(data);
    
    //从 Excel 文件中选择第一个工作表（Sheet）。
    let sheet=
    workbook.Sheets[
        workbook.SheetNames[0]
    ];


    let result=
    XLSX.utils.sheet_to_json(sheet);
    console.log(result);
    
    words =
    result.map(item=>({
        
        word:String(item.word).trim(),
        
        note:String(item.note).trim()
        
    }))
    .filter(item=>item.word!="undefined" && item.note!="undefined");



    //显示词库数量
    
    document
    .getElementById("word-count-info")
    .innerHTML =
    "词库数量：" + words.length + " 个";


});





//读取错题词库

fetch('/data/wrong_words.xlsx')

.then(response=>response.arrayBuffer())

.then(data=>{


    let workbook=XLSX.read(data);


    let sheet=
    workbook.Sheets[
        workbook.SheetNames[0]
    ];


    let result=
    XLSX.utils.sheet_to_json(sheet);



wrongWords =
[
    ...new Map(
        result
        .map(item=>[
            item.word,
            {
                word:item.word,
                note:item.note
            }
        ])
    ).values()
];
    


});


//点击按钮随机显示 


function test(){

    if(words.length===0){
    
        alert("普通词库正在加载，请稍等");
    
        return;
    
    }


    //读取随机种子

    let seedValue =
    document
    .getElementById("seed")
    .value;



    if(seedValue){

        seed =
        Number(seedValue);

    }
    else{

        //没有输入种子，使用当前时间

        seed =
        Date.now()%100000;

    }   
    // 获取用户需要数量
    
    let count =
    Number(
        document.getElementById("word-count").value
    );



    // 理论需要错题数量（40%）
    
    let targetWrong = Math.ceil(count * 0.4);


    // 实际错题数量
    
    let wrongCount;
     
    if(wrongWords.length < targetWrong){
    
        // 错题不足40%，全部显示
    
        wrongCount = wrongWords.length;
    
    }
    
    else{
    
        // 错题足够，只取40%
    
        wrongCount = targetWrong;
    
    }



    // 普通词补足数量
    
    let normalCount = count - wrongCount;




    // ==================
    // 随机错题
    // ==================
    
    let wrongTemp =
    shuffle(
        [...wrongWords]
    );


    let selectedWrong =
    wrongTemp.slice(
        0,
        wrongCount
    );





    // ==================
    // 随机普通词
    // ==================
    
    let normalTemp =
    shuffle(
        [...words]
    );


    let selectedNormal =
    normalTemp.slice(
        0,
        normalCount
    );





    // 合并
    
    let selected =
    [
        ...selectedWrong,
        ...selectedNormal
    ];




    // 最后整体打乱
    
    selected =
    shuffle(selected);
    // 去除同一个单词的中英文重复
    let usedWords = new Set();
    let usedNotes = new Set();


    selected = selected.filter(item=>{


        if(
            usedWords.has(item.word) ||
            usedNotes.has(item.note)
        ){

            return false;

        }


        usedWords.add(item.word);

        usedNotes.add(item.note);


        return true;


    });



     


currentWords = selected;


document
.getElementById("word-list")
.innerHTML =


        selected
        .map((item,index)=>{


            let number=index+1;


            // 随机显示英文或者中文

            let showEnglish = seededRandom()<0.5;


            let question;


            if(showEnglish){

                question=item.word;

            }
            else{

                question=item.note;

            }



            return `

            <div class="word-card">


            <div class="question">

            ${number}. ${question}

            </div>

       


            <input

            class="answer-input"

            id="answer-${index}"

            placeholder="请输入答案">



            <div

            class="result"

            id="result-${index}">

            </div>


            </div>


            `;


        })
        .join(""); 

}
 
function showAnswer(){


    currentWords.forEach((item,index)=>{


        let userAnswer =
        document
        .getElementById(
            "answer-"+index
        )
        .value
        .trim();



        let result =
        document
        .getElementById(
            "result-"+index
        );



        let answer;


        // 判断输入方向

        let question =
        document
        .querySelectorAll(".question")
        [index]
        .innerText;

        answer=item.note;

        if(question.includes(item.word)){


            //显示英文，需要回答中文

            answer=item.note;


            result.innerHTML= answer;


        }

        else{


            //显示中文，需要回答英文

            answer=item.word;

            result.innerHTML= answer;


        }
        result.innerHTML =      
        `
        <button 
        class="speak-btn"
        onclick="speak('${item.word}')">
        ▶️
        </button>
        `
        + answer ;


    });


}
function seededRandom(){


    seed =
    (seed * 9301 + 49297) % 233280;


    return seed / 233280;


}
function shuffle(array){


    for(
        let i=array.length-1;
        i>0;
        i--
    ){


        let j =
        Math.floor(
            seededRandom()*(i+1)
        );



        [
            array[i],
            array[j]
        ] =
        [
            array[j],
            array[i]
        ];


    }


    return array;

}
function speak(word){


    let speech =
    new SpeechSynthesisUtterance(word);


    speech.lang="en-US";


    speech.rate= 1;


    speech.pitch=1;


    window.speechSynthesis.speak(speech);


}


let totalSeconds=25*60;
let startSeconds=25*60;
function updateTimer(){


    let min=Math.floor(totalSeconds/60);

    let sec=totalSeconds%60;


    document.getElementById("timer").innerHTML=

    min+":"+
    String(sec).padStart(2,"0");



    let percent =
    totalSeconds/startSeconds*100;


    document
    .getElementById("progress-bar")
    .style.width =
    percent+"%";


}
let timer = null;


//开始计时

function startTimer(){


    if(timer){

        return;

    }


    timer = setInterval(()=>{


        if(totalSeconds > 0){


            totalSeconds--;

            updateTimer();


        }
        else{


            clearInterval(timer);

            timer=null;

            alert("🍅 番茄时间结束！");

        }


    },1000);


}



//暂停

function pauseTimer(){


    clearInterval(timer);

    timer=null;


}



//重置

function resetTimer(){


    clearInterval(timer);

    timer=null;


    totalSeconds=startSeconds;


    updateTimer();


}

function goTop(){

    window.scrollTo({

        top:0,

        behavior:"smooth"

    });

}
</script>

<button onclick="showAnswer()">
显示正确答案
</button>
<button id="top-btn" onclick="goTop()">
    <span></span>
</button>

