import random

class MarkovChain:
    def __init__(self, transition_matrix):
        self.transition_matrix = transition_matrix
        self.num_states = len(transition_matrix)

    def get_transition_probability(self, current_state, new_state):
        return self.transition_matrix[current_state - 1][new_state - 1]

    def simulate(self, initial_state, num_steps):
        state_counts = [0] * self.num_states
        current_state = initial_state

        for _ in range(num_steps):
            r = random.random()
            cumulative_probability = 0

            for new_state in range(1, self.num_states + 1):
                cumulative_probability += self.get_transition_probability(current_state, new_state)

                if r <= cumulative_probability:
                    state_counts[new_state - 1] += 1
                    current_state = new_state
                    break

        probabilities = [count / num_steps for count in state_counts]
        return probabilities

def print_probabilities(probabilities):
    print(f"p1 = {probabilities[0]:.3f}, p2 = {probabilities[1]:.3f}, p3 = {probabilities[2]:.3f}, p4 = {probabilities[3]:.3f}")

def main():
    transition_matrix = [
        [0, 0, 0.8, 0.2],  
        [0, 0.1, 0.5, 0.4],      
        [0.2, 0.2, 0.2, 0.4],          
        [0, 0, 0, 1],    
    ]

    markov_chain = MarkovChain(transition_matrix)
    initial_state = 2
    num_steps = 1000

if __name__ == "__main__":
    main()








