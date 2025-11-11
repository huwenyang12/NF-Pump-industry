"""
给定一个整数数组 nums 和一个整数目标值 target，请你在该数组中找出 和为目标值 target  的那 两个 整数，并返回它们的数组下标。
"""

from typing import List


class Solution:

    def __init__(self):
        self.nums = [3,7,11,6]
        self.target = 9
        pass

    def twoSum_1(self, nums: List[int], target: int) -> List[int]:
        for i in range(len(nums)):
            for j in range(i+1,len(nums)):
                if nums[i] + nums[j] == target:
                    return [i ,j]
                


if __name__ == "__main__":
    obj = Solution()
    obj.twoSum_1()


# 