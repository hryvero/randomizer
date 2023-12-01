## Overview

This repository contains an Express API . It is designed to authenticate users.

## Table of Contents

- [Getting Started](#getting-started)
    - [Prerequisites](#prerequisites)
    - [Installation](#installation)
- [Usage](#usage)




## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) installed
- [npm](https://www.npmjs.com/) or [Yarn](https://yarnpkg.com/) package manager

### Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/hryvero/randomizer.git

2. Install dependencies:
   Copy code

    ```bash
    cd randomizer
    npm install


### PreStart
Fill in .env file with credentials described in .env.example
Copy code

    npm start


#  Usage
You can use it manually or automatically.

# Manually
/start

Then you should upload Excel file with  these fields: 

Topic(String),

Speaker(String),

Status (true or false), 

Username (Telegram username)

/generate 

For generating random speaker

# Automatically

Just set time in cronTime and wait for response.