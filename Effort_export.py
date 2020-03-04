import unittest

def eat(food, is_healthy):
    if is_healthy:
        return "I am healthy"
    else:
        return "I do not eat because it is unhealthy"


def nap(num_hours):
    if num_hours < 8:
        return 'I need to sleep more'
    if num_hours > 8:
        return 'Sleep well, live well'


def play(game, hours):
    if (len(game)) > 0:
        if hours > 3:
            return "I play more than 3 hours"
        else:
            return "I play less or equal than 3 hours"
    else:
        return "I do not play games"
 
class SomeTests(unittest.TestCase):
    def test_play(self):
        """testing a thing"""
        self.assertEqual(play('Fifa', 4), "I play more than 3 hours")

    def test_eat(self):
        """testing another thing"""
        self.assertEqual(eat('Banana', is_healthy=True), "I am healthy")

    def test_nap(self):
        """testing another thing"""
        self.assertEqual(nap(9), 'Sleep well, live well')
        
if __name__ == "__main__":
    unittest.main()
    
