# Logical Data Model

Core entities:
- Portfolio
- Broker
- Security
- Transaction
- Dividend
- Watchlist
- PriceHistory
- ExchangeRate
- Configuration

Relationships:
Portfolio 1:N Transaction
Security 1:N PriceHistory
Broker 1:N Transaction
