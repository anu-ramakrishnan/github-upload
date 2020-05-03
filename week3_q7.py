"""
Implement the Sieve of Eratosthenes
https://en.wikipedia.org/wiki/Sieve_of_Eratosthenes
"""

def compute_primes(bound):
    """
    Return a list of the prime numbers in range(2, bound)
    """
    answer = list(range(2, bound))
    for divisor in range(2, bound):
        # Remove appropriate multiples of divisor from answer
        for i in range(divisor,len(answer)-1):            
            if(answer[i]%divisor == 0):
                answer[i] = 0
    #print(answer)
    length = len(answer)
    i=0
    while(i<length):
        if (answer[i] == 0):
            answer.remove(answer[i])
            length=length-1
            continue
        i+=1
    #print(answer)
    return answer
               
    
print(len(compute_primes(200)))
print(len(compute_primes(2000)))