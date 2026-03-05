#!/bin/bash

# Docker deployment script for BAF PDF Batch Processor
# Usage: ./deploy.sh [command]
# Example: ./deploy.sh deploy

set -e  # Exit on any error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Configuration
APP_NAME="baf-pdf-processor"
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
COMPOSE_FILE="${APP_DIR}/docker-compose.yml"

# Logging functions
log() {
    echo -e "${GREEN}[$(date +'%Y-%m-%d %H:%M:%S')]${NC} $1"
}

error() {
    echo -e "${RED}[ERROR]${NC} $1" >&2
    exit 1
}

warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

# Check Docker is installed and running
check_docker() {
    log "Checking Docker..."
    if ! command -v docker &> /dev/null; then
        error "Docker is not installed. Please install Docker and try again."
    fi
    if ! docker info &> /dev/null; then
        error "Docker daemon is not running. Please start Docker and try again."
    fi
    log "Docker check passed ✓"

    log "Checking Docker Compose..."
    if ! docker compose version &> /dev/null 2>&1; then
        if ! command -v docker-compose &> /dev/null; then
            error "Docker Compose is not installed. Please install Docker Compose and try again."
        fi
        COMPOSE_CMD="docker-compose"
    else
        COMPOSE_CMD="docker compose"
    fi
    log "Docker Compose check passed ✓"
    export COMPOSE_CMD
}

# Setup directories for volume mounts
setup_directories() {
    log "Setting up required directories..."
    local dirs=("${APP_DIR}/input" "${APP_DIR}/output")

    if ! mkdir -p "${dirs[@]}" 2>/dev/null; then
        # Directories may exist with root ownership (from previous Docker run)
        log "Fixing directory permissions (may require sudo)..."
        sudo mkdir -p "${dirs[@]}"
        sudo chown -R "$(whoami):" "${APP_DIR}/input" "${APP_DIR}/output"
    fi

    if [[ "$OSTYPE" != "msys" && "$OSTYPE" != "win32" ]]; then
        chmod 755 "${APP_DIR}/input" "${APP_DIR}/output" 2>/dev/null || true
    fi
    log "Directories ready ✓"
}

# Check .env file
check_env() {
    if [ ! -f "${APP_DIR}/.env" ]; then
        warning ".env file not found. Application may run with default/empty configuration."
        if [ ! -f "${APP_DIR}/.env.example" ]; then
            cat > "${APP_DIR}/.env.example" << 'EOF'
# OpenAI Configuration
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-4

# Application Configuration
LOG_LEVEL=INFO
EOF
            warning "Created .env.example. Copy to .env and configure before production use."
        fi
    else
        log "Environment file found ✓"
    fi
}

# Start scheduler (runs in background, monitors input folder)
deploy() {
    log "Starting BAF PDF scheduler..."
    log "Application directory: $APP_DIR"

    check_docker
    setup_directories
    check_env

    log "Building and starting scheduler..."
    cd "$APP_DIR"
    $COMPOSE_CMD -f "$COMPOSE_FILE" up -d --build

    log "Scheduler started. Place PDFs in ./input - they will be processed automatically."
    log "Output Excel files will appear in ./output"
    log ""
    log "Useful commands:"
    log "  View logs:    $COMPOSE_CMD -f $COMPOSE_FILE logs -f"
    log "  Stop:        ./deploy.sh stop"
    log "  Status:      ./deploy.sh status"
}

# Build image only (no run)
build() {
    log "Building Docker image..."
    check_docker
    setup_directories
    cd "$APP_DIR"
    $COMPOSE_CMD -f "$COMPOSE_FILE" build
    log "Build complete. Run: ./deploy.sh deploy to start"
}

# Stop scheduler
stop() {
    log "Stopping scheduler..."
    check_docker
    cd "$APP_DIR"
    $COMPOSE_CMD -f "$COMPOSE_FILE" down
    log "Scheduler stopped ✓"
}

# Show status
status() {
    check_docker
    cd "$APP_DIR"
    log "Container status:"
    $COMPOSE_CMD -f "$COMPOSE_FILE" ps
}

# Remove stopped containers
clean() {
    log "Cleaning up Docker resources..."
    check_docker
    cd "$APP_DIR"
    $COMPOSE_CMD -f "$COMPOSE_FILE" down
    log "Cleanup complete ✓"
}

# Main script logic
case "${1:-deploy}" in
    deploy)
        deploy
        ;;
    build)
        build
        ;;
    stop)
        stop
        ;;
    status)
        status
        ;;
    logs)
        check_docker
        cd "$APP_DIR"
        $COMPOSE_CMD -f "$COMPOSE_FILE" logs -f
        ;;
    clean)
        clean
        ;;
    run-once)
        check_docker
        setup_directories
        cd "$APP_DIR"
        $COMPOSE_CMD -f "$COMPOSE_FILE" run --rm app python process_pdfs.py
        ;;
    *)
        echo "Usage: $0 {deploy|build|stop|status|logs|clean|run-once}"
        echo ""
        echo "Commands:"
        echo "  deploy    - Build and start scheduler (monitors input folder)"
        echo "  build     - Build Docker image only"
        echo "  stop      - Stop the scheduler"
        echo "  status    - Show container status"
        echo "  logs      - Follow scheduler logs"
        echo "  clean     - Remove Docker resources"
        echo "  run-once  - Run one-time batch (no scheduler)"
        echo ""
        echo "Scheduler monitors ./input for new PDFs and writes Excel to ./output"
        exit 1
        ;;
esac
