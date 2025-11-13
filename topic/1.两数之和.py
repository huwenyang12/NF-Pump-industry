"""
给定一个整数数组 nums 和一个整数目标值 target，请你在该数组中找出 和为目标值 target  的那 两个 整数，并返回它们的数组下标。
"""

from typing import List


class Solution:

    def __init__(self):
        pass

    def twoSum_1(self, nums: List[int], target: int) -> List[int]:
        for i in range(len(nums)):
            for j in range(i+1,len(nums)):
                if nums[i] + nums[j] == target:
                    return[i, j]
    
    def twoSum_2(self, nums: List[int], target: int) -> List[int]:
        data = {}
        for idx,value in enumerate(nums):
            a = target - value
            if a in data:
                return[data[a],idx]
            data[value] = idx
        return []



if __name__ == "__main__":
    nums = [3,7,11,6]
    target = 9

    obj = Solution()
    list = obj.twoSum_2(nums,target)

    print(list)


# 