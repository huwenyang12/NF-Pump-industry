class Solution:
    def __init__(self):
        pass

    def isPalindrome(self, x: int) -> bool:
        return str(x) == str(x)[::-1]


if __name__ == "__main__":
    x = int(input("请输入数字："))
    obj = Solution()
    print(obj.isPalindrome(x))