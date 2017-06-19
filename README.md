# MonteCarlo-Excel

This macro prompts the user for the specified worksheet, initial stock price, expected rate of return, dividend yield rate, volatility, time steps, number of simulations, and time period length. It then takes the inputs and runs Monte Carlo simulations, outputting the prices to the cells of the specified worksheet in a matrix format, with each column representing a simulation. The macro then dynamically generates a 2-D line chart that plots the stock price paths against each other. Stock prices are assumed to follow a geometric Brownian motion.

Next steps will be to extend the macro to price European and exotic options, incorporate variance reduction methods, and implement stochastic volatility models.
