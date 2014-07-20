
# Motivation

This class was written to work around limitations of the standard Collection class:

- find() returns a reference to existing TreeNode -> its payload can be changed.
- can do partial iteration (For Each with Collection object can only do full; iteration with Collection.Item is O(N^2)).
- duplicate keys don't have to be treated as errors.
- add() can use simple flat array as key (as long as array members can be compared using '<' and '=').
- any object can be key as long as TreeMap.cmp_key() is adapted.

# Classes and Methods

## TreeMap
	- add(key As Variant, value As Variant) As TreeNode
	- find(key As Variant) As TreeNode
	- count() as Long
	- remove(key As Variant)
	- inorder(Optional from_key As Variant) As TreeInorderCursor
	- dump(Optional N As TreeNode)

## TreeNode

All members are public for simplicity.

	- payload as Variant

## TreeInorderCursor
	- next_node() As TreeNode
	- prev_node() As TreeNode

Methods next_ a prev_ will return Nothing when the Map is exhausted.

	- first() As TreeNode
	- last() As TreeNode
	- start(start_at As TreeNode)
