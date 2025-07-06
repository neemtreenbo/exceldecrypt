# Use official Node.js image
FROM node:18

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy source code
COPY . .

# Expose port (if you use something like Express.js)
EXPOSE 3000

# Start the app
CMD ["npm", "start"]