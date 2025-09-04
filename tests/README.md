# Microsoft MCP Tests

This directory contains unit tests for the Microsoft Graph MCP (Model Context Protocol) server.

## Test Structure

### Core Test Files

- **`test_auth.py`** - Tests for Azure authentication module
  - Authentication credential creation and configuration
  - Token management and caching
  - AuthenticationRecord persistence
  - Graph client creation
  - Environment variable handling

- **`test_graph.py`** - Tests for Microsoft Graph API interaction module
  - HTTP request handling and retries
  - Pagination support
  - Search query functionality
  - Error handling for various HTTP status codes
  - Rate limiting and exponential backoff

- **`test_tools_simple.py`** - Tests for MCP tools core functionality
  - Folder mapping configuration
  - Parameter validation logic
  - Search and pagination parameter construction
  - Endpoint URL construction logic

- **`test_integration.py`** - Integration tests between modules
  - Module imports and dependencies
  - Auth and graph module interaction
  - Configuration validation
  - Logging setup verification

### Test Configuration

- **`conftest.py`** - Shared test fixtures and configuration
  - Mock authentication instances
  - Sample data fixtures (users, emails)
  - Environment cleanup utilities

## Running Tests

```bash
# Run all tests
pytest tests/

# Run tests with verbose output
pytest tests/ -v

# Run specific test file
pytest tests/test_auth.py -v

# Run tests with coverage (if coverage package is installed)
pytest tests/ --cov=src/microsoft_mcp
```

## Test Coverage

The tests cover the following key areas:

### Authentication (`test_auth.py`)
- ✅ Azure credential creation with various configurations
- ✅ Token acquisition and caching mechanisms
- ✅ AuthenticationRecord serialization/deserialization
- ✅ Environment variable validation
- ✅ Error handling for missing credentials
- ✅ Graph client instantiation

### Graph API (`test_graph.py`)
- ✅ HTTP request construction and execution
- ✅ Authentication header injection
- ✅ Search request special headers (consistency level, prefer)
- ✅ Rate limiting (429) and server error (5xx) retry logic
- ✅ Pagination handling with @odata.nextLink
- ✅ Search query with entity type filtering
- ✅ Various HTTP error code handling

### Tools Logic (`test_tools_simple.py`)
- ✅ Folder name mapping and case-insensitive lookup
- ✅ Email parameter validation and limits
- ✅ Search parameter construction
- ✅ Endpoint URL building logic
- ✅ Availability check payload creation
- ✅ Dependency injection verification

### Integration (`test_integration.py`)
- ✅ Module import verification
- ✅ Cross-module communication (auth ↔ graph)
- ✅ Configuration constants
- ✅ Base URL and endpoint validation

## Test Strategy

The tests focus on:

1. **Unit Testing**: Individual functions and classes in isolation
2. **Logic Testing**: Core business logic without external dependencies
3. **Integration Testing**: Module interactions and dependency injection
4. **Error Handling**: Various failure scenarios and edge cases
5. **Configuration Testing**: Environment setup and parameter validation

## Limitations

Due to the FastMCP decorator pattern used in the tools module, the tests focus on:
- Testing the underlying logic and dependencies
- Validating configuration and constants
- Ensuring proper module imports and setup
- Testing core business logic components

Direct function execution testing is handled through dependency mocking and logic validation rather than calling decorated functions directly.

## Dependencies

The tests use:
- `pytest` - Test framework
- `unittest.mock` - Mocking and patching
- Standard library modules for test utilities

All test dependencies are managed through the `[dependency-groups.dev]` section in `pyproject.toml`.
