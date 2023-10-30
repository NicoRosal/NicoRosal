let blockSize = 25,
    rows = 20,
    cols = 20,
    // snake head
    snakeX = blockSize * 5,
    snakeY = blockSize * 5,
    velocityX = 0,
    velocityY = 0;
    //snake body
    snakeBody = [],
    // food
    foodX = blockSize * 10,
    foodY = blockSize * 10,
    gameOver = false;

var board,
    context;

window.onload = function() {
    board = document.getElementById("board");
    board.height = rows * blockSize;
    board.width = rows * blockSize;
    context = board.getContext("2d"); //Used for drawing on board

    placeFood();
    document.addEventListener("keyup", changeDirection)
    //update();
    setInterval(update, 1000/10);
}

function update() {
    if (gameOver) {
        return;
    }

    context.fillStyle="black";
    context.fillRect(0, 0, board.width, board.height);

    context.fillStyle="white";
    context.fillRect(foodX, foodY, blockSize, blockSize);

    if (snakeX == foodX && snakeY == foodY) {
        snakeBody.push([foodX, foodY])
        placeFood();
    }

    for (let i = snakeBody.length-1; i > 0; i--) {
        snakeBody[i] = snakeBody[i-1];
    }

    if (snakeBody && snakeBody.length) {
        snakeBody[0] = [snakeX, snakeY];
    }

    context.fillStyle="green";
    snakeX += velocityX * blockSize;
    snakeY += velocityY * blockSize;
    context.fillRect(snakeX, snakeY, blockSize, blockSize);

    for (let i = 0; i< snakeBody.length; i++) {
        context.fillRect(snakeBody[i][0], snakeBody[i][1], blockSize, blockSize);
    }

    //game over conditions
    if (snakeX < 0 || snakeX > cols*blockSize-1 || snakeY < 0 || snakeY > rows*blockSize-1) {
        gameOver = true;
        alert("Game Over");
    }

    for (let i = 0; i < snakeBody.length; i++) {
        if (snakeX == snakeBody[i][0] && snakeY == snakeBody[i][1]) {
            gameOver = true;
            alert("Game Over");
        }
    }
}

function changeDirection(e) {
    if (e.code == "ArrowUp" && velocityY != 1) {
        velocityX = 0;
        velocityY = -1;
    }
    if (e.code == "ArrowDown" && velocityY != -1) {
        velocityX = 0;
        velocityY = 1;
    }
    if (e.code == "ArrowLeft" && velocityX != 1) {
        velocityX = -1;
        velocityY = 0;
    }
    if (e.code == "ArrowRight" && velocityX != -1) {
        velocityX = 1;
        velocityY = 0;
    }    
}

function placeFood() {
    //(0-1) * cols -> (0-19.999) -> (0-19) * 25
    foodX = Math.floor(Math.random() * cols) * blockSize;
    foodY = Math.floor(Math.random() * rows) * blockSize;
}