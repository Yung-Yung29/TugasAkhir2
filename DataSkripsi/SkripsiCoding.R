library(plotly)
library(reshape2)
library(openxlsx)
library(ggplot2)
library(dplyr)
library("writexl")
library(openxlsx)

# Parameters
alpha = 1/90 # External infection rate
beta = 1/30  # Internal infection rate
gamma = 1/50    # Quarantined rate
kappa = 1/100  # Recovery rate
tao <- 1/300   # Transition rate to Permanently Damaged state

# Number of computers, iterations, simulations, and interest rate
N = 10
n_simulations = 100
n_iterations = 100
int_rate = 0.02

# Starting loop for simulations with different seeds
premiums_data <- list()
state_durations_data <- list()  
num_transitions_data <- list()
process_data <- list()
total_time <- system.time({
  for (simul in 1:n_simulations) {
    for (iter in 1:n_iterations) {
      # Start time for each iteration and simulation
      # Initialize matrices to track state durations and number of transitions
      state_durations <- matrix(0, nrow = N, ncol = 4)  # N computers, 4 states
      rownames(state_durations) = c(paste("N", 1:N))
      colnames(state_durations) = c(1:4)
      num_transitions <- matrix(0, nrow = 5, ncol = N)
      colnames(num_transitions) = c(paste("N", 1:N))
      
      # Initialize matrix to store and track computer states
      process = data.frame(matrix(ncol = N + 5))  # Increased by 1 for Permanently Damaged
      colnames(process) = c("t", paste("Com", 1:N), "n_1", "n_2", "n_3", "n_4")
      
      tic <- Sys.time()
      
      # Initialize the first time for simulation: 0 
      current.t = 0
      process[1, 1] = current.t
      
      # Initialize vectors for S, I, Q, P
      S = rep(0, N)
      I = rep(0, N)
      Q = rep(0, N)
      P = rep(0, N)
      
      # Initialize initial status for every computer and the number of computers in that state
      process[1, 2:(1 + N)] = 1 
      process[1, N + 2] = sum(process[1, 2:(N + 1)] == 1)
      process[1, N + 3] = 0  # Initial number of computers in Infected state is 0
      process[1, N + 4] = 0  # Initial number of computers in Quarantined state is 0
      process[1, N + 5] = 0  # Initial number of computers in Permanently Damaged state is 0
      
      # Initialize the first infection  
      init.inf = round(rexp(N, alpha), 6)  # Generate time-to-failure external infection
      init.inf.t = min(init.inf)  # Find minimum time
      process[2, 1] = init.inf.t  # Make that as the next-time for the first infection
      
      # Find and set which computer with minimum time to failure 
      susceptible_comp <- setdiff(1:N, which.min(init.inf))
      infected_comp <- which.min(init.inf)
      
      # Change the status of the infected computer
      process[2, 1 + susceptible_comp] = 1
      process[2, 1 + infected_comp] = 2
      
      # Make an initial status for S, I, Q, P: 0
      S[susceptible_comp] = 1
      I[infected_comp] = 1
      
      # Count the number of S, I, Q, P
      process[2, N + 2] = sum(process[2, 2:(N + 1)] == 1)
      process[2, N + 3] = sum(process[2, 2:(N + 1)] == 2)
      process[2, N + 4] = sum(process[2, 2:(N + 1)] == 3)
      process[2, N + 5] = sum(process[2, 2:(N + 1)] == 4)
      
      # Make the current time as the next step for simulation
      current.t = init.inf.t
      
      # Update state durations and the number of infected computers
      num_transitions[1, infected_comp] <- num_transitions[1, infected_comp] + 1
      
      i = 2  # NextRow of process data frame
      
      while (current.t <= 365) {
        # Generate NA for the number of computers
        inf_t <- rep(NA, N)
        qua_t <- rep(NA, N)
        rec_t <- rep(NA, N)
        dam2_t <- rep(NA, N)
        dam3_t <- rep(NA, N)
        
        # Build a matrix that contains all possible times
        time = as.data.frame(matrix(ncol = N))
        
        # Generate time-to-failure
        for (comp in which(S == 1)) {
          inf_ex = round(rexp(1, alpha), 6)
          
          # Check if there are any infected computers
          if (any(I == 1)) {
            inf_in = round(rexp(1, beta), 6)
          } else {
            inf_in = numeric(0)  # Empty vector
          }
          # Find the minimum time from external and infection time
          inf_t[comp] = min(c(inf_ex, inf_in))
        }
        
        # Generate time-to-quarantined
        for (comp in which(I == 1)) {
          # Check if there are any infected computers
          if (any(I == 1)) {
            qua_t[comp] = round(rexp(1, gamma), 6)
            dam2_t[comp] = round(rexp(1, tao), 6)
          }
        }
        
        # Generate time-to-recovered
        for (comp in which(Q == 1)) {
          rec_t[comp] = round(rexp(1, kappa), 6)
          dam3_t[comp] = round(rexp(1, tao), 6)
        }
        
        # Combine all of the time-to vectors into the time dataframe
        time = as.data.frame(rbind(inf_t, qua_t, rec_t, dam2_t, dam3_t))  # Added Quarantined and Permanently Damaged
        
        # Rename the time dataframe headers
        for (j in 1:N) {
          colnames(time)[j] = paste("Comp", j, sep = "")
        }
        
        # Locate the minimum time and type of event
        for (j in 1:(dim(time)[1])) {
          min_values <- na.omit(as.numeric(time[j, 1:N]))
          if (length(min_values) > 0) {
            time$min.t[j] = min(min_values)
          } else {
            # Handle the case where all values are missing
            time$min.t[j] = Inf  # Set to a large value 
          }
        }
        
        index_min = which.min(time$min.t)  # Index which event occurred
        loc.comp = as.numeric(which.min(time[index_min, 1:N]))  # Comp which changed state
        
        if (current.t + time$min.t[index_min] <= 365) {
          process[i + 1, 1] = process[i, 1] + time$min.t[index_min]
          if (rownames(time)[index_min] == "inf_t") {
            process[i + 1, loc.comp + 1] = 2
            num_transitions[1, loc.comp] <- num_transitions[1, loc.comp] + 1
            S[loc.comp] = 0
            I[loc.comp] = 1
          } else if (rownames(time)[index_min] == "qua_t") {
            process[i + 1, loc.comp + 1] = 3  # Change from 2 to 3
            num_transitions[2, loc.comp] <- num_transitions[2, loc.comp] + 1
            I[loc.comp] = 0
            Q[loc.comp] = 1  # Increment Quarantined count
          } else if (rownames(time)[index_min] == "rec_t") {
            process[i + 1, loc.comp + 1] = 1  # Change from 3 to 1
            num_transitions[3, loc.comp] <- num_transitions[3, loc.comp] + 1
            Q[loc.comp] = 0
            S[loc.comp] = 1
          } else if (rownames(time)[index_min] == "dam2_t") {
            if (process[i, loc.comp + 1] == 2) { # Adjust this line
              process[i + 1, loc.comp + 1] = 4 # Change from 2 to 4
              num_transitions[4, loc.comp] <- num_transitions[4, loc.comp] + 1
              I[loc.comp] = 0
              P[loc.comp] = 1 # Set computer to Permanently Damaged state
            }
          } else if (rownames(time)[index_min] == "dam3_t") {
            if (process[i, loc.comp + 1] == 3) { # Adjust this line
              process[i + 1, loc.comp + 1] = 4 # Change from 3 to 4
              num_transitions[5, loc.comp] <- num_transitions[5, loc.comp] + 1
              Q[loc.comp] = 0
              P[loc.comp] = 1 # Set computer to Permanently Damaged state
            }
          }
          
          process[i + 1, -c(1, loc.comp + 1, (N + 2):(N + 5))] = process[i, -c(1, loc.comp + 1, (N + 2):(N + 5))]
          
          # Count the number of S, I, Q, P
          process[i + 1, N + 2] = sum(process[i + 1, 2:(N + 1)] == 1)
          process[i + 1, N + 3] = sum(process[i + 1, 2:(N + 1)] == 2)
          process[i + 1, N + 4] = sum(process[i + 1, 2:(N + 1)] == 3)  # Count Quarantined
          process[i + 1, N + 5] = sum(process[i + 1, 2:(N + 1)] == 4)  # Count Permanently Damaged
          
          current.t = process[i + 1, 1]
          i = i + 1
        } else {
          
          i = i - 1 
          
          process_2 = process
          process_2$Infected_Increment[1] = 0
          process_2$Quarantined_Increment[1] = 0  # Initialize Quarantined Increment
          process_2$PermanentlyDamaged_Increment[1] = 0  # Initialize Permanently Damaged Increment
          
          for (j in 2:dim(process_2)[1]) {
            t = process_2[j, 1]
            if (!is.na(process_2$n_2[j - 1])) {
              process_2$Infected_Increment[j] = process_2$n_2[j] - process_2$n_2[j - 1]
            } else {
              process_2$Quarantined_Increment[j] = 0
            }
            if (!is.na(process_2$n_3[j - 1])) {
              process_2$Quarantined_Increment[j] = process_2$n_3[j] - process_2$n_3[j - 1]
            } else {
              process_2$Quarantined_Increment[j] = 0
            }
          }
          
          pv_premium <- 0
          for (row in 2:dim(process_2)[1]) {
            if (process_2$Quarantined_Increment[row] == 1) {
              time2 <- process_2[row, 1]
              pv_premium <- pv_premium + (1 + int_rate)^(-time2 / 365)
            }
          }
          
          # Store PV premium in the list
          premiums_data[[length(premiums_data) + 1]] <- list(Simulation = simul, Iteration = iter, PV_Premium = pv_premium)
          
          if (any(process[, 1] <= 365)) {
            # Iterate through each computer
            for (comp in 1:N) {
              # Initialize variables to track state and start time
              current_state <- process[1, comp + 1]
              start_time <- process[1, 1]
              
              # Iterate through each row in 'process'
              for (k in 2:(nrow(process)-1)) {
                time3 <- process[k, 1]
                new_state <- process[k, comp + 1]
                
                # If the state has changed
                if (new_state != current_state) {
                  # Update the state duration matrix
                  state_durations[comp, current_state] <- state_durations[comp, current_state] + (time3 - start_time)
                  
                  # Update current state and start time
                  current_state <- new_state
                  start_time <- time3
                }
              }
              
              # If the last state persists until the end
              if (current_state != 0) {
                state_durations[comp, current_state] <- state_durations[comp, current_state] + (365 - start_time)
              }
            }
            
            # Store durations data in the list
            state_durations_data[[length(state_durations_data) + 1]] <- list(Simulation = simul, Iteration = iter, StateDuration = state_durations)
            
            #Store num_transitions data in the list
            num_transitions_data[[length(num_transitions_data) + 1]] <- list(Simulation = simul, Iteration = iter, NumTransitions = num_transitions)
          }
          
          process_data[[length(process_data) + 1]] <- process
          toc <- Sys.time()
          
          elapsed_time <- difftime(toc, tic, units = "secs")
          print(paste("Elapsed Time (secs) for Simulation", simul, "Iteration", iter, ":", elapsed_time))
          break
        }
      }
    }
  }
})

premiums_info <- lapply(premiums_data, function(x) c(x$Simulation, x$Iteration, x$PV_Premium))
premiums_info <- do.call(rbind, premiums_info)
colnames(premiums_info) <- c("Simulation", "Iteration", "PV_Premium")
premiums_info <- as.data.frame(premiums_info)
average_premiums <- premiums_info %>%
  group_by(Simulation) %>%
  summarise(Average_PV_Premium = mean(PV_Premium))

total_transitions <- matrix(0, nrow = 5, ncol = 10)

# Loop through num_transitions_data
for (i in 1:length(num_transitions_data)) {
  transitions <- num_transitions_data[[i]]$NumTransitions
  total_transitions <- total_transitions + transitions
}

# Calculate average transitions
average_transitions <- total_transitions / length(num_transitions_data)

for (i in 1:length(state_durations_data)) {
  # Transpose the StateDuration matrix
  state_durations <- t(state_durations_data[[i]]$StateDuration)
  
  # Assign the transposed matrix back to StateDuration
  state_durations_data[[i]]$StateDuration <- state_durations
}

total_durations <- matrix(0, nrow = 4, ncol = 10)

# Loop through state_durations_data
for (i in 1:length(state_durations_data)) {
  # Extract state durations for each simulation and iteration
  durations <- state_durations_data[[i]]$StateDuration
  
  # Sum the durations for each state across all simulations and iterations
  total_durations <- total_durations + durations
}

# Calculate the average duration for each state
average_durations <- total_durations / length(state_durations_data)

# Inizialition matrix that has infinite value
infinite_matrices_indices <- vector("list", length = length(state_durations_data))
count <- 1

# Loop through state_durations_data
for (i in 1:length(state_durations_data)) {
  matrix <- state_durations_data[[i]]$StateDuration
  
  # Check if there is infinite or not?
  if (any(is.infinite(matrix))) {
    infinite_matrices_indices[[count]] <- i
    count <- count + 1
  }
}

negative_elements_state_durations <- sapply(state_durations_data, function(data) any(unlist(data$StateDuration) < 0))
# Identify which simulations and iterations have negative elements
simulations_with_negative_state_durations <- which(negative_elements_state_durations)

negative_elements_num_transitions <- sapply(num_transitions_data, function(data) any(unlist(data$Num_Transitions) < 0))
simulations_with_negative_num_transitions <- which(negative_elements_state_durations)

# Extract the average premiums from the data frame
Average_premiums <- average_premiums$Average_PV_Premium

# Shapiro-Wilk test for normality
shapiro_test <- shapiro.test(Average_premiums)

# Print the result
print(shapiro_test)

ks_test <- ks.test(Average_premiums, "pnorm", mean = mean(Average_premiums), sd = sd(Average_premiums))

# Print the result
print(ks_test)

qqnorm(average_premiums$Average_PV_Premium, main = "Q-Q Plot of Average Premiums")
qqline(average_premiums$Average_PV_Premium, col = 2)

# Add labels
xlabel <- "Theoretical Quantiles"
ylabel <- "Sample Quantiles"
mtext(xlabel, side=1, line=3)
mtext(ylabel, side=2, line=3)

mean_premium <- mean(average_premiums$Average_PV_Premium)
sd_premium <- sd(average_premiums$Average_PV_Premium)

# Generate a sequence of values along the x-axis
x <- seq(min(average_premiums$Average_PV_Premium), max(average_premiums$Average_PV_Premium), length=100)

# Generate the normal distribution curve
y <- dnorm(x, mean = mean_premium, sd = sd_premium)

# Create a histogram
hist(average_premiums$Average_PV_Premium, probability=TRUE, col="lightblue", breaks=20, main="Histogram with Bell Curve")
lines(x, y, col="red", lwd=2)

# Add legend
legend("topright", legend=c("Data", "Normal Curve"), col=c("lightblue", "red"), lwd=1)

average_premiums_sorted<- average_premiums[order(average_premiums$Average_PV_Premium, decreasing = FALSE), ]
confidence_level <- 0.95
var_95 <- quantile(average_premiums_sorted$Average_PV_Premium, 1 - confidence_level)

num_obs <- nrow(average_premiums_sorted)
index <- floor(confidence_level * num_obs)

# Calculate the TVaR (average of values beyond the threshold)
tvar <- mean(average_premiums_sorted$Average_PV_Premium[(index+1):num_obs])

if (any(negative_elements_num_transitions)) {
  print("Ada nilai TRUE di dalam vektor.")
} else {
  print("Tidak ada nilai TRUE di dalam vektor.")
}

if (any(negative_elements_state_durations)) {
  print("Ada nilai TRUE di dalam vektor.")
} else {
  print("Tidak ada nilai TRUE di dalam vektor.")
}

if (any(sapply(infinite_matrices_indices, function(x) !is.null(x)))) {
  print("Ada elemen yang tidak NULL dalam daftar.")
} else {
  print("Semua elemen dalam daftar adalah NULL.")
}

z_scores <- (average_premiums$Average_PV_Premium - mean_premium) / sd_premium

# Generate a sequence of values along the x-axis
x <- seq(min(z_scores), max(z_scores), length=100)

# Generate the standard normal distribution curve
y <- dnorm(x, mean = 0, sd = 1)

# Create a histogram
hist(z_scores, probability=TRUE, col="lightblue", breaks=20, main="Histogram with Standardized Normal Curve")
lines(x, y, col="red", lwd=2)

# Labeling axes
xlabel <- "Standardized Values"
ylabel <- "Density"

mean = mean(average_premiums$Average_PV_Premium)
sd = sd(average_premiums$Average_PV_Premium)
CI_lower = mean-(1.96*(sd/10))
CI_upper = mean+(1.96*(sd/10))

ggplot(average_premiums, aes(x = Simulation, y = Average_PV_Premium)) +
  geom_line() +
  geom_ribbon(aes(ymin = CI_lower, ymax = CI_upper), alpha = 0.3, fill = "blue") +
  geom_point() +
  labs(title = "Grafik Rata-Rata dengan Confidence Interval") +
  theme_minimal()
outliers <- subset(average_premiums, Average_PV_Premium < CI_lower | Average_PV_Premium > CI_upper)
num_outliers <- nrow(outliers)
t_test_additional <- t.test(Average_PV_Premium ~ ifelse(Average_PV_Premium < CI_lower | Average_PV_Premium > CI_upper, "Outside CI", "Inside CI"), data = average_premiums)

split_num_transitions_data <- split(num_transitions_data, sapply(num_transitions_data, `[[`, "Simulation"))
average_transitions_per_simulation <- lapply(split_num_transitions_data, function(sim_data) {
  if (length(sim_data) > 0) {
    return(do.call(rbind, lapply(sim_data, function(matrix_data) {
      if (nrow(matrix_data$NumTransitions) > 0) {
        return(matrix(colMeans(matrix_data$NumTransitions, na.rm = TRUE), ncol = 10, byrow = TRUE))
      } else {
        return(matrix(rep(NA, 10), ncol = 10, byrow = TRUE))  # or any value that indicates missing data
      }
    })))
  } else {
    return(NULL)  # or any value that indicates missing data
  }
})

calculate_average_matrix <- function(sim_data) {
  return(sapply(1:5, function(row_index) {
    rowMeans(do.call(cbind, lapply(sim_data, function(matrix_data) {
      if (nrow(matrix_data$NumTransitions) > 0) {
        return(matrix_data$NumTransitions[row_index, , drop = TRUE])
      } else {
        return(rep(NA, length(sim_data)))  
      }
    })), na.rm = TRUE)
  }))
}

all_average_matrices <- lapply(split_num_transitions_data, calculate_average_matrix)
all_average_matrices <- lapply(all_average_matrices, function(matrix_data) t(matrix_data))
average_matrices_per_simulation <- lapply(all_average_matrices, function(matrix_data) {
  if (nrow(matrix_data) > 0) {
    return(data.frame(t(rowMeans(matrix_data, na.rm = TRUE))))
  } else {
    return(NULL)  
  }
})
transposed_average_matrices <- lapply(average_matrices_per_simulation, function(df) {
  if (!is.null(df)) {
    return(t(df))
  } else {
    return(NULL)  
  }
})

list_of_dataframes <- list()


for (row_index in seq_len(5)) {
  row_values <- lapply(transposed_average_matrices, function(matrix_data) {
    if (!is.null(matrix_data)) {
      return(matrix_data[row_index, , drop = FALSE])
    } else {
      return(NULL)  
    }
  })
  
  combined_values <- do.call(cbind, row_values)
  
  new_dataframe <- as.data.frame(combined_values)
  
  list_of_dataframes[[row_index]] <- new_dataframe
}

transposed_dataframes <- lapply(list_of_dataframes, function(df) {
  if (!is.null(df)) {
    return(t(df))
  } else {
    return(NULL)  
  }
})

first_list_values <- transposed_dataframes[[1]]
second_list_values <- transposed_dataframes[[2]]
third_list_values <- transposed_dataframes[[3]]
fourth_list_values <- transposed_dataframes[[4]]
fifth_list_values <- transposed_dataframes[[5]]

correlation <- cor(average_premiums$Average_PV_Premium, first_list_values, use = "complete.obs")
correlation2 <- cor(average_premiums$Average_PV_Premium, second_list_values, use = "complete.obs")
correlation3 <- cor(average_premiums$Average_PV_Premium, third_list_values, use = "complete.obs")
correlation4 <- cor(average_premiums$Average_PV_Premium, fourth_list_values, use = "complete.obs")
correlation5 <- cor(average_premiums$Average_PV_Premium, fifth_list_values, use = "complete.obs")

par(mfrow = c(2, 3))

plot(average_premiums$Average_PV_Premium, first_list_values, main = "Correlation", 
     xlab = "Rata-Rata Premi", ylab = "Transisi 1->2", pch = 16, col = "blue")
abline(lm(first_list_values ~ average_premiums$Average_PV_Premium), col = "red")

plot(average_premiums$Average_PV_Premium, second_list_values, main = "Correlation", 
     xlab = "Rata-Rata Premi", ylab = "Transisi 2-> 3", pch = 16, col = "blue")
abline(lm(second_list_values ~ average_premiums$Average_PV_Premium), col = "red")

plot(average_premiums$Average_PV_Premium, third_list_values, main = "Correlation", 
     xlab = "Rata-Rata Premi", ylab = "Transisi 3-> 1", pch = 16, col = "blue")
abline(lm(third_list_values ~ average_premiums$Average_PV_Premium), col = "red")

plot(average_premiums$Average_PV_Premium, fourth_list_values, main = "Correlation", 
     xlab = "Rata-Rata Premi", ylab = "Transisi 2-> 4", pch = 16, col = "blue")
abline(lm(fourth_list_values ~ average_premiums$Average_PV_Premium), col = "red")

plot(average_premiums$Average_PV_Premium, fifth_list_values, main = "Correlation", 
     xlab = "Rata-Rata Premi", ylab = "Transisi 3-> 4", pch = 16, col = "blue")
abline(lm(fifth_list_values ~ average_premiums$Average_PV_Premium), col = "red")


premi <- c(15.75252, 15.85188, 15.93825, 14.65607, 15.22583, 14.72267)
parameter <- c("Kontrol", "Alpha", "Beta", "Gamma", "Kappa", "Tau")
data <- data.frame(Parameter = parameter, Premi = premi)

# Plot line menggunakan ggplot2
ggplot(data, aes(x = parameter, y = Premi, group = 1)) +
  geom_line() +
  geom_point() +
  labs(x = "Parameter", y = "Rata-RataPremi") +
  theme_minimal()

stdpremi = c(0.3560548, 0.340398, 0.3359988, 0.3137707, 0.2827653, 0.3202078)
parameter2 <- c("Kontrol", "Alpha", "Beta", "Gamma", "Kappa", "Tau")
data2 <- data.frame(Parameter2 = parameter2, STDPremi = stdpremi)

ggplot(data2, aes(x = parameter2, y = stdpremi, group = 1)) +
  geom_line() +
  geom_point() +
  labs(x = "Parameter", y = "Standar Deviasi Premi") +
  theme_minimal()

data <- data.frame(
  Parameter = c("Kontrol", "Alpha", "Beta", "Gamma", "Kappa", "Tau"),
  EP = c(15.75252, 15.85188, 14.65607, 15.93825, 15.22583, 14.72267),
  VaR = c(15.24386, 15.35404, 14.16815, 15.3086, 14.8512, 14.1194),
  TVaR = c(16.5559, 16.53717, 15.27961, 16.56583, 15.82728, 15.36796)
)

# Reshape data for better plotting
data_long <- tidyr::gather(data, Key, Value, -Parameter)

# Plot
ggplot(data_long, aes(x = Parameter, y = Value, group = Key, color = Key)) +
  geom_line() +
  geom_point() +
  labs(title = "Perbandingan Premi Bersih Menggunakan EP, VaR, dan TVaR",
       x = "Parameter", y = "Nilai Premi Bersih") +
  theme_minimal()