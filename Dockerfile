FROM node:20-slim

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install all dependencies (including devDependencies for build)
RUN npm ci

# Copy source code
COPY src/ ./src/
COPY public/ ./public/
COPY tsconfig.json ./

# Build TypeScript
RUN npm run build

# Remove devDependencies after build
RUN npm prune --production

# Create necessary directories
RUN mkdir -p uploads work output

# Expose port (Zeabur will use PORT env var)
EXPOSE 3000

# Run the application
CMD ["node", "dist/server.js"]
