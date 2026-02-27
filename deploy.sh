#!/bin/bash

# Deployment script for FastAPI Document Extraction Application
# Usage: ./deploy.sh [environment]
# Example: ./deploy.sh production

set -e  # Exit on any error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Configuration
ENVIRONMENT="${1:-production}"
APP_NAME="document-extraction-app"
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${APP_DIR}/venv"
LOG_DIR="${APP_DIR}/logs"
PID_FILE="${APP_DIR}/app.pid"
PORT="${PORT:-7005}"
HOST="${HOST:-0.0.0.0}"
WORKERS="${WORKERS:-4}"

# Logging function
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

# Check if running as root (not recommended for production)
if [ "$EUID" -eq 0 ]; then
    warning "Running as root is not recommended for production deployments"
fi

# Check Python version
check_python() {
    log "Checking Python version..."
    
    # Try python3 first, then python (for Windows/Git Bash compatibility)
    if command -v python3 &> /dev/null; then
        PYTHON_CMD="python3"
    elif command -v python &> /dev/null; then
        PYTHON_CMD="python"
    else
        error "Python 3 is not installed. Please install Python 3.8 or higher."
    fi
    
    PYTHON_VERSION=$($PYTHON_CMD --version 2>&1 | cut -d' ' -f2 | cut -d'.' -f1,2)
    REQUIRED_VERSION="3.8"
    
    if [ "$(printf '%s\n' "$REQUIRED_VERSION" "$PYTHON_VERSION" | sort -V | head -n1)" != "$REQUIRED_VERSION" ]; then
        error "Python 3.8 or higher is required. Found: $PYTHON_VERSION"
    fi
    
    log "Python version check passed: $($PYTHON_CMD --version)"
    export PYTHON_CMD
}

# Create virtual environment
setup_venv() {
    log "Setting up virtual environment..."
    
    if [ ! -d "$VENV_DIR" ]; then
        log "Creating virtual environment..."
        $PYTHON_CMD -m venv "$VENV_DIR"
    else
        log "Virtual environment already exists"
    fi
    
    log "Activating virtual environment..."
    # Handle both Unix and Windows paths
    if [ -f "${VENV_DIR}/bin/activate" ]; then
        source "${VENV_DIR}/bin/activate"
    elif [ -f "${VENV_DIR}/Scripts/activate" ]; then
        source "${VENV_DIR}/Scripts/activate"
    else
        error "Could not find virtual environment activation script"
    fi
    
    log "Upgrading pip..."
    pip install --upgrade pip setuptools wheel --quiet
}

# Install dependencies
install_dependencies() {
    log "Installing dependencies..."
    
    if [ ! -f "${APP_DIR}/requirements.txt" ]; then
        error "requirements.txt not found!"
    fi
    
    pip install -r "${APP_DIR}/requirements.txt" --quiet
    log "Dependencies installed successfully"
}

# Setup directories
setup_directories() {
    log "Setting up required directories..."
    
    # Create storage/uploads directory
    mkdir -p "${APP_DIR}/storage/uploads"
    chmod 755 "${APP_DIR}/storage/uploads"
    
    # Create logs directory
    mkdir -p "$LOG_DIR"
    chmod 755 "$LOG_DIR"
    
    # Create .gitkeep if needed
    touch "${APP_DIR}/storage/uploads/.gitkeep"
    
    log "Directories created successfully"
}

# Check environment variables
check_env() {
    log "Checking environment variables..."
    
    if [ ! -f "${APP_DIR}/.env" ]; then
        warning ".env file not found. Creating template..."
        cat > "${APP_DIR}/.env.example" << EOF
# OpenAI Configuration
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-4

# Application Configuration
LOG_LEVEL=INFO
PORT=7005
HOST=0.0.0.0

# Server Configuration
WORKERS=4
EOF
        warning "Please create .env file with your configuration. A template has been created at .env.example"
    else
        log "Environment file found"
    fi
    
    # Load environment variables
    if [ -f "${APP_DIR}/.env" ]; then
        set -a
        source "${APP_DIR}/.env"
        set +a
    fi
}

# Stop existing application
stop_app() {
    log "Checking for running application..."
    
    if [ -f "$PID_FILE" ]; then
        OLD_PID=$(cat "$PID_FILE")
        if ps -p "$OLD_PID" > /dev/null 2>&1; then
            log "Stopping existing application (PID: $OLD_PID)..."
            kill "$OLD_PID" || true
            sleep 2
            
            # Force kill if still running
            if ps -p "$OLD_PID" > /dev/null 2>&1; then
                warning "Force killing process..."
                kill -9 "$OLD_PID" || true
            fi
        fi
        rm -f "$PID_FILE"
        log "Application stopped"
    else
        log "No running application found"
    fi
}

# Start application
start_app() {
    log "Starting application..."
    
    # Activate virtual environment
    source "${VENV_DIR}/bin/activate"
    
    # Load environment variables
    if [ -f "${APP_DIR}/.env" ]; then
        set -a
        source "${APP_DIR}/.env"
        set +a
    fi
    
    # Override with script variables if set
    export PORT="${PORT:-7005}"
    export HOST="${HOST:-0.0.0.0}"
    export WORKERS="${WORKERS:-4}"
    
    log "Starting uvicorn server on ${HOST}:${PORT} with ${WORKERS} workers..."
    
    # Start uvicorn in background
    nohup uvicorn app.main:app \
        --host "$HOST" \
        --port "$PORT" \
        --workers "$WORKERS" \
        --log-level "${LOG_LEVEL:-info}" \
        --access-log \
        --log-config "${APP_DIR}/logging.conf" 2>/dev/null || \
    uvicorn app.main:app \
        --host "$HOST" \
        --port "$PORT" \
        --workers "$WORKERS" \
        --log-level "${LOG_LEVEL:-info}" \
        --access-log \
        > "${LOG_DIR}/app.log" 2>&1 &
    
    APP_PID=$!
    echo $APP_PID > "$PID_FILE"
    
    # Wait a moment and check if process is still running
    sleep 2
    if ps -p "$APP_PID" > /dev/null 2>&1; then
        log "Application started successfully (PID: $APP_PID)"
        log "Server running on http://${HOST}:${PORT}"
        log "Logs are available at ${LOG_DIR}/app.log"
    else
        error "Failed to start application. Check logs at ${LOG_DIR}/app.log"
    fi
}

# Health check
health_check() {
    log "Performing health check..."
    
    sleep 3
    
    if curl -f -s "http://${HOST}:${PORT}/" > /dev/null 2>&1; then
        log "Health check passed ✓"
    else
        warning "Health check failed. Application may still be starting..."
    fi
}

# Main deployment function
deploy() {
    log "Starting deployment for environment: $ENVIRONMENT"
    log "Application directory: $APP_DIR"
    
    check_python
    setup_venv
    install_dependencies
    setup_directories
    check_env
    stop_app
    start_app
    health_check
    
    log "Deployment completed successfully!"
    log ""
    log "Useful commands:"
    log "  - View logs: tail -f ${LOG_DIR}/app.log"
    log "  - Stop app: ./deploy.sh stop"
    log "  - Restart app: ./deploy.sh restart"
}

# Stop function
stop() {
    log "Stopping application..."
    stop_app
    log "Application stopped"
}

# Restart function
restart() {
    log "Restarting application..."
    stop_app
    sleep 2
    start_app
    health_check
    log "Application restarted"
}

# Status function
status() {
    if [ -f "$PID_FILE" ]; then
        PID=$(cat "$PID_FILE")
        if ps -p "$PID" > /dev/null 2>&1; then
            log "Application is running (PID: $PID)"
            log "Port: ${PORT:-7005}"
            log "Host: ${HOST:-0.0.0.0}"
        else
            warning "PID file exists but process is not running"
            rm -f "$PID_FILE"
        fi
    else
        warning "Application is not running"
    fi
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
    *)
        echo "Usage: $0 {deploy|stop|restart|status} [environment]"
        echo ""
        echo "Commands:"
        echo "  deploy   - Deploy/update the application (default)"
        echo "  stop     - Stop the running application"
        echo "  restart  - Restart the application"
        echo "  status   - Check application status"
        echo ""
        echo "Examples:"
        echo "  $0 deploy production"
        echo "  $0 restart"
        echo "  $0 stop"
        exit 1
        ;;
esac
