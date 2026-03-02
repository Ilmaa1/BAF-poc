#!/bin/bash

# Docker deployment script for FastAPI Document Extraction Application
# Usage: ./deploy.sh [command]
# Example: ./deploy.sh deploy

set -e  # Exit on any error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Configuration
APP_NAME="document-extraction-app"
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
COMPOSE_FILE="${APP_DIR}/docker-compose.yml"
PORT="${PORT:-7005}"

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
    mkdir -p "${APP_DIR}/storage/uploads" "${APP_DIR}/storage/records" "${APP_DIR}/logs"
    if [[ "$OSTYPE" != "msys" && "$OSTYPE" != "win32" ]]; then
        chmod 755 "${APP_DIR}/storage/uploads" "${APP_DIR}/storage/records" "${APP_DIR}/logs"
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

# Deploy application
deploy() {
    log "Starting Docker deployment..."
    log "Application directory: $APP_DIR"

    check_docker
    setup_directories
    check_env

    log "Building and starting containers..."
    cd "$APP_DIR"
    $COMPOSE_CMD -f "$COMPOSE_FILE" up -d --build

    log "Deployment completed successfully!"
    log ""
    log "Application running at: http://localhost:${PORT}"
    log ""
    log "Useful commands:"
    log "  - View logs:    $COMPOSE_CMD -f $COMPOSE_FILE logs -f"
    log "  - Stop app:    ./deploy.sh stop"
    log "  - Restart app: ./deploy.sh restart"
    log "  - Status:      ./deploy.sh status"
}

# Stop application
stop() {
    log "Stopping Docker containers..."
    check_docker
    cd "$APP_DIR"
    $COMPOSE_CMD -f "$COMPOSE_FILE" down
    log "Application stopped ✓"
}

# Restart application
restart() {
    log "Restarting Docker containers..."
    check_docker
    cd "$APP_DIR"
    $COMPOSE_CMD -f "$COMPOSE_FILE" restart
    log "Application restarted ✓"
}

# Show status
status() {
    check_docker
    cd "$APP_DIR"
    log "Container status:"
    $COMPOSE_CMD -f "$COMPOSE_FILE" ps
}

# Main script logic
case "${1:-deploy}" in
    deploy)
        deploy
        ;;
    stop)
        stop
        ;;
    restart)
        restart
        ;;
    status)
        status
        ;;
    logs)
        check_docker
        cd "$APP_DIR"
        $COMPOSE_CMD -f "$COMPOSE_FILE" logs -f
        ;;
    *)
        echo "Usage: $0 {deploy|stop|restart|status|logs}"
        echo ""
        echo "Commands:"
        echo "  deploy   - Build and start the application (default)"
        echo "  stop     - Stop the running containers"
        echo "  restart  - Restart the containers"
        echo "  status   - Show container status"
        echo "  logs     - Follow container logs"
        echo ""
        echo "Examples:"
        echo "  $0 deploy"
        echo "  $0 stop"
        echo "  $0 restart"
        exit 1
        ;;
esac
