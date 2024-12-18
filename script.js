let questions = [];
let currentQuestion = 0;
let score = 0;
let userAnswers = [];

// 開始測驗
async function startTest() {
    try {
        // 載入所有題目
        const response = await fetch('questions.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);
        
        // 處理題目數據
        questions = jsonData.map(row => {
            // 檢查是否為選擇題（通過是否有選項來判斷）
            if (row.option1 && row.option2) {
                return {
                    type: 'multiple-choice',
                    question: row.question,
                    options: [row.option1, row.option2, row.option3, row.option4].filter(Boolean), // 過濾掉空值
                    correctAnswer: parseInt(row.correctAnswer) - 1
                };
            } else {
                return {
                    type: 'fill-in',
                    question: row.question,
                    correctAnswer: row.answer.toString() // 確保答案為字串
                };
            }
        });

        console.log('Loaded questions:', questions); // 用於調試
        shuffleQuestions();
        document.getElementById('start-section').style.display = 'none';
        document.getElementById('quiz-section').style.display = 'block';
        showQuestion();
    } catch (error) {
        alert('載入題目時發生錯誤，請稍後再試');
        console.error('Error loading questions:', error);
    }
}

// 顯示當前題目
function showQuestion() {
    const questionType = document.getElementById('question-type');
    const questionText = document.getElementById('question-text');
    const answerContainer = document.getElementById('answer-container');
    
    // 顯示題型
    const currentType = questions[currentQuestion].type;
    questionType.className = `question-type ${currentType}`;
    questionType.textContent = currentType === 'multiple-choice' ? '選擇題' : '填充題';
    
    questionText.textContent = `${currentQuestion + 1}. ${questions[currentQuestion].question}`;
    answerContainer.innerHTML = '';
    
    if (currentType === 'multiple-choice') {
        questions[currentQuestion].options.forEach((option, index) => {
            const optionDiv = document.createElement('div');
            optionDiv.className = 'option';
            optionDiv.textContent = option;
            if (userAnswers[currentQuestion] === index) {
                optionDiv.classList.add('selected');
            }
            optionDiv.onclick = () => selectOption(index);
            answerContainer.appendChild(optionDiv);
        });
    } else {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'fill-in-input';
        input.placeholder = '請輸入答案';
        if (userAnswers[currentQuestion]) {
            input.value = userAnswers[currentQuestion];
        }
        input.oninput = (e) => {
            userAnswers[currentQuestion] = e.target.value;
        };
        answerContainer.appendChild(input);
    }

    updateNavigationButtons();
    updateProgressBar();
}

// 更新導航按鈕
function updateNavigationButtons() {
    document.getElementById('next-btn').style.display = 
        currentQuestion === questions.length - 1 ? 'none' : 'block';
    document.getElementById('submit-btn').style.display = 
        currentQuestion === questions.length - 1 ? 'block' : 'none';
}

// 更新進度條
function updateProgressBar() {
    const progress = ((currentQuestion + 1) / questions.length) * 100;
    document.getElementById('progress-bar').style.width = `${progress}%`;
}

// 選擇選項（選擇題）
function selectOption(index) {
    const options = document.querySelectorAll('.option');
    options.forEach(option => option.classList.remove('selected'));
    options[index].classList.add('selected');
    userAnswers[currentQuestion] = index;
}

// 下一題
function nextQuestion() {
    if (questions[currentQuestion].type === 'fill-in' && 
        (!userAnswers[currentQuestion] || userAnswers[currentQuestion].trim() === '')) {
        if (!confirm('您尚未填寫答案，確定要繼續下一題嗎？')) {
            return;
        }
    }
    
    if (currentQuestion < questions.length - 1) {
        currentQuestion++;
        showQuestion();
    }
}

// 檢查是否所有題目都已作答
function checkAllAnswered() {
    const unanswered = questions.reduce((acc, question, index) => {
        const answer = userAnswers[index];
        if (answer === undefined || answer.toString().trim() === '') {
            acc.push(index + 1);
        }
        return acc;
    }, []);

    if (unanswered.length > 0) {
        return confirm(`第 ${unanswered.join(', ')} 題尚未作答，確定要提交嗎？`);
    }
    return true;
}

// 修改提交函數
function submitQuiz() {
    if (!checkAllAnswered()) {
        return;
    }

    score = 0;
    const reviewContainer = document.getElementById('answer-review');
    reviewContainer.innerHTML = '<h3>答題檢討：</h3>';

    questions.forEach((question, index) => {
        const userAnswer = userAnswers[index];
        let isCorrect = false;

        if (question.type === 'multiple-choice') {
            isCorrect = userAnswer === question.correctAnswer;
        } else {
            // 填充題答案比對：移除多餘空格並轉為小寫進行比較
            const cleanUserAnswer = (userAnswer || '').toString().trim().toLowerCase();
            const cleanCorrectAnswer = question.correctAnswer.toString().trim().toLowerCase();
            isCorrect = cleanUserAnswer === cleanCorrectAnswer;
        }

        if (isCorrect) score++;

        // 添加檢討內容
        const reviewItem = document.createElement('div');
        reviewItem.className = `review-item ${isCorrect ? 'correct' : 'incorrect'}`;
        
        if (question.type === 'multiple-choice') {
            reviewItem.innerHTML = `
                <p><strong>題目 ${index + 1} (選擇題):</strong> ${question.question}</p>
                <p>您的答案: ${userAnswer !== undefined ? question.options[userAnswer] : '未作答'}</p>
                <p>正確答案: ${question.options[question.correctAnswer]}</p>
            `;
        } else {
            reviewItem.innerHTML = `
                <p><strong>題目 ${index + 1} (填充題):</strong> ${question.question}</p>
                <p>您的答案: ${userAnswer || '未作答'}</p>
                <p>正確答案: ${question.correctAnswer}</p>
            `;
        }
        
        reviewContainer.appendChild(reviewItem);
    });

    // 顯示結果
    document.getElementById('quiz-section').style.display = 'none';
    document.getElementById('result-section').style.display = 'block';
    
    // 計算百分比
    const percentage = (score / questions.length * 100).toFixed(1);
    document.getElementById('score').innerHTML = `
        <div class="score-detail">
            <p>總題數: ${questions.length} 題</p>
            <p>答對題數: ${score} 題</p>
            <p>正確率: ${percentage}%</p>
        </div>
    `;
}

// 返回選擇頁面
function backToSelection() {
    currentQuestion = 0;
    score = 0;
    userAnswers = [];
    questions = [];
    
    document.getElementById('result-section').style.display = 'none';
    document.getElementById('test-selection').style.display = 'block';
}

// 打亂題目順序
function shuffleQuestions() {
    for (let i = questions.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [questions[i], questions[j]] = [questions[j], questions[i]];
    }
}