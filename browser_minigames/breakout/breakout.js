let board,
    boardWidth = 500,
    boardHeight = 500,
    context,
    //player
    playerWidth = 80,
    playerHeight = 10,
    playerVelocityX = 10,
    //ball
    ballWidth = 10,
    ballHeight = 10,
    ballVelocityX = 3,
    ballVelocityY = 2,
    //blcoks
    blockArray = [],
    blockWidth = 50,
    blockHeight = 10,
    blockColumns = 8,
    blockRows = 3,
    blockMaxRows = 10,
    blockCount = 0,
//starting block corner top left
    blockX = 15,
    blockY = 45,
    score = 0,
    gameOver = false;

let ball = {
    x: boardWidth/2,
    y : boardHeight/2,
    width : ballWidth,
    height : ballHeight,
    velocityX : ballVelocityX,
    velocityY: ballVelocityY
}

let player = {
        x : boardWidth/2 - playerWidth/2,
        y : boardHeight - playerHeight - 5,
        width : playerWidth,
        height : playerHeight,
        velocityX : playerVelocityX
    }

window.onload = function() {
    board = document.getElementById("board");
    board.height = boardHeight;
    board.width = boardWidth;
    context = board.getContext("2d");

    //draw player
    context.fillStyle = "yellow";
    context.fillRect(player.x, player.y, player.width, player.height);

    requestAnimationFrame(update);
    document.addEventListener("keydown", moveplayer);

    createBlocks();
}

function update() {
    requestAnimationFrame(update);
    context.clearRect(0, 0, board.width, board.height);

    //player
    context.fillStyle = "yellow";
    context.fillRect(player.x, player.y, player.width, player.height);

    //ball
    context.fillStyle = "white";
    ball.x += ball.velocityX;
    ball.y += ball.velocityY;
    context.fillRect(ball.x, ball.y, ball.width, ball.height);

    //bounce logic
    if (topCollision(ball, player) || bottomCollision(ball, player)) {
        ball.velocityY *= -1; // flip direction up or down
    } else if (leftCollision(ball, player) || rightCollision(ball, player)) {
        ball.velocityX *= -1;
    }
    
    //bounce ball
    if (ball.y <= 0) {
        //if ball touch top of canvas
        ball.velocityY *= -1;
    } else if (ball.x <= 0 || (ball.x + ball.width) >= boardWidth) {
        //ball touch left or right border
        ball.velocityX *= -1;
    } else if ((ball.y + ball.height) >= boardHeight) {
        // if ball touches the ground, game over
        context.font = "20px sans-serif";
        context.fillText("Game Over: Please press 'Space' to Restart", 80, 400);
        gameOver = true;
    }

    //blocks
    context.fillStyle = "blue";
    for (let i = 0; i < blockArray.length; i++) {
        let block = blockArray[i];
        if (!block.break) {
            if (topCollision(ball, block) || bottomCollision(ball, block)) {
                block.break = true;
                ball.velocityY *= -1;
                score += 100;
                blockCount -= 1;
            } else if (leftCollision(ball, block) || rightCollision(ball, block)) {
                block.break = true;
                ball.velocityX *= -1;
                score += 100;
                blockCount -= 1;
            }
            context.fillRect(block.x, block.y, block.width, block.height);
        }
    }
    
    //next level
    if (blockCount == 0) {
        score += 100*blockRows*blockColumns;
        blockRows = Math.min(blockRows + 1, blockMaxRows);
        createBlocks();
    }

    //score
    context.font = "20px sans-serif";
    context.fillText(score, 10, 25);
}

function outOfBounds(xPosition) {
    return (xPosition < 0 || (xPosition + playerWidth) > boardWidth)
}
function moveplayer(e) {
    if (gameOver) {
        if (e.code = "Space") {
            resetGame();
            console.log("RESET");
        }
        return;
    }
    if (e.code == "ArrowLeft") {
        //player.x -= player.velocityX;
        let nextPlayerX = player.x - player.velocityX; 
        
        if (!outOfBounds(nextPlayerX)) {
            player.x = nextPlayerX;
        }
    } else if (e.code == "ArrowRight") {
        //player.x += player.velocityX;
        let nextPlayerX = player.x + player.velocityX;
 
        if (!outOfBounds(nextPlayerX)) {
            player.x = nextPlayerX;
        }
    }
}

function detectCollision(a, b) {
    return a.x < b.x + b.width && //a's top left corner doesn't reach b's top right corer
           a.x + a.width > b.x && //a's top right corner passes b's top left corner
           a.y < b.y + b.height && //a's top left corener doesn't reach b's bottom left corner
           a.y + a.height > b.y; //a's bottom left corner passes b's top left corner
}

function topCollision(ball, block) { //ball is above block
    return detectCollision(ball, block) && (ball.y + ball.height) >= block.y;
}

function bottomCollision(ball, block) {//ball is below block
    return detectCollision(ball, block) && (block.y + block.height) >= ball.y;
}

function leftCollision(ball, block) {
    return detectCollision(ball, block) && (ball.x + ball.width) >= block.x;
}

function rightCollision(ball, block) {
    return detectCollision(ball, block) && (block.x + block.width) >= ball.x;
}

function createBlocks() {
    blockArray = [];
    for (let i =0; i < blockColumns; i++) {
        for (let j =0; j < blockColumns; j++) {
            let block = {
                x : blockX + i*blockWidth + i*10,
                y : blockY + j*blockHeight + j*10,
                width : blockWidth,
                height : blockHeight,
                break : false
            }
            blockArray.push(block);
        }
    }
    blockCount = blockArray.length;
}

function resetGame() {
    gameOver = false;
    player = { 
        x : boardWidth/2 - playerWidth/2,
        y: boardHeight - playerHeight - 5,
        width: playerWidth,
        height: playerHeight,
        velocityX : playerVelocityX
    }
    ball = {
        x : boardWidth/2,
        y: boardHeight/2,
        width: ballWidth,
        height: ballHeight,
        velocityX : ballVelocityX,
        velocityY : ballVelocityY
    }
    blockArray = [];
    blockRows = 3;
    score = 0;
    createBlocks();
}