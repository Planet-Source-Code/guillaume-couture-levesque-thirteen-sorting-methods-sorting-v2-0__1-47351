Attribute VB_Name = "modSort"
'This code is the property of Guillaume Couture-Levesque
'Do not use this code in any other program without the
'permission of the author (me(Guillaume Couture-Levesque))

Option Explicit
Option Base 0

Public Sub InsertionSort(numbers() As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim index As Integer
    Dim end_now As Boolean
    
    'do the insertion sort
    For i = 1 To (num - 1)
        end_now = False
        index = numbers(i)
        j = i
        k = j - 1
        Do While ((j > 0) And (Not end_now))
            If (numbers(k) > index) Then
                numbers(j) = numbers(k)
                j = j - 1
                k = j - 1
            Else
                end_now = True
            End If
        Loop
        numbers(j) = index
    Next i
End Sub

Public Sub SelectionSort(numbers() As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim min As Integer
    Dim temp As Integer
    
    'do the selection sort
    For i = 0 To (num - 2)
        min = i
        For j = (i + 1) To (num - 1)
            If numbers(j) < numbers(min) Then
                min = j
            End If
        Next j
        temp = numbers(i)
        numbers(i) = numbers(min)
        numbers(min) = temp
    Next i
End Sub

Public Sub BubbleSort(numbers() As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    Dim swapped As Boolean
    
    'do the bubble sort
    For i = (num - 1) To 0 Step -1
        swapped = False
        For j = 1 To i
            If (numbers(j - 1) > numbers(j)) Then
                swapped = True
                temp = numbers(j - 1)
                numbers(j - 1) = numbers(j)
                numbers(j) = temp
            End If
        Next j
        If Not swapped Then
            i = -1
        End If
    Next i
End Sub

Public Sub ShakerSort(numbers() As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp As Integer
    Dim swapped As Boolean
    
    'do the shaker sort aka double-bouble aka cocktail aka bi-directional bubble
    For k = 1 To ((num - 1) \ 2) + 1
        'set the initial swapped
        swapped = False
        
        'up
        For i = k To (num - k)
            If numbers(i - 1) > numbers(i) Then
                swapped = True
                temp = numbers(i - 1)
                numbers(i - 1) = numbers(i)
                numbers(i) = temp
            End If
        Next i
        
        'down
        For j = (num - k - 1) To (k - 1) Step -1
            If numbers(j + 1) < numbers(j) Then
                swapped = True
                temp = numbers(j + 1)
                numbers(j + 1) = numbers(j)
                numbers(j) = temp
            End If
        Next j
        
        'check for early exit
        If Not swapped Then
            k = num
        End If
    Next k
End Sub

Public Sub ShellSort(numbers() As Integer, num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim increment As Integer
    Dim temp As Integer
    Dim end_now As Boolean
    
    'do the shell sort
    increment = 3
    Do While (increment > 0)
        For i = 0 To (num - 1)
            end_now = False
            j = i
            k = j - increment
            temp = numbers(i)
            Do While ((j >= increment) And (Not end_now))
                If (numbers(k) > temp) Then
                    numbers(j) = numbers(k)
                    j = k
                    k = k - increment
                Else
                    end_now = True
                End If
            Loop
            numbers(j) = temp
        Next i
        If (increment \ 2 <> 0) Then
            increment = increment \ 2
        ElseIf (increment = 1) Then
            increment = 0
        Else
            increment = 1
        End If
    Loop
End Sub

Public Sub HeapSort(numbers() As Integer, num As Integer)
    Dim i As Integer
    Dim temp As Integer
    
    'create the heap
    For i = ((num \ 2) - 1) To 0 Step -1
        siftDown numbers, i, num
    Next i
    
    'sort the heap
    For i = (num - 1) To 1 Step -1
        temp = numbers(0)
        numbers(0) = numbers(i)
        numbers(i) = temp
        siftDown numbers, 0, i - 1
    Next i
End Sub

Public Sub siftDown(numbers() As Integer, root As Integer, bottom As Integer)
    Dim done As Integer
    Dim maxChild As Integer
    Dim temp As Integer
    
    'creating the heap
    done = 0
    Do While ((root * 2 <= bottom) And (done <> 1))
        If (root * 2 = bottom) Then
            maxChild = root * 2
        ElseIf (numbers(root * 2) > numbers(root * 2 + 1)) Then
            maxChild = root * 2
        Else
            maxChild = root * 2 + 1
        End If
        
        If (numbers(root) < numbers(maxChild)) Then
            temp = numbers(root)
            numbers(root) = numbers(maxChild)
            numbers(maxChild) = temp
            root = maxChild
        Else
            done = 1
        End If
    Loop
End Sub

Public Sub MergeSort(numbers() As Integer, temp() As Integer, ByVal num As Integer)
    'do the merge sort
    m_sort numbers, temp, 0, num - 1
End Sub

Public Sub m_sort(numbers() As Integer, temp() As Integer, ByVal left As Integer, ByVal right As Integer)
    Dim mid As Integer
    
    'do the recursive aspect
    If (right > left) Then
        mid = (right + left) \ 2
        m_sort numbers, temp, left, mid
        m_sort numbers, temp, mid + 1, right
        
        merge numbers, temp, left, mid + 1, right
    End If
End Sub

Public Sub merge(numbers() As Integer, temp() As Integer, ByVal left As Integer, ByVal mid As Integer, ByVal right As Integer)
    Dim i As Integer
    Dim left_end As Integer
    Dim num_elements As Integer
    Dim tmp_pos As Integer
    
    'merge the correct elements
    left_end = mid - 1
    tmp_pos = left
    num_elements = right - left + 1
    
    Do While ((left <= left_end) And (mid <= right))
        If (numbers(left) <= numbers(mid)) Then
            temp(tmp_pos) = numbers(left)
            tmp_pos = tmp_pos + 1
            left = left + 1
        Else
            temp(tmp_pos) = numbers(mid)
            tmp_pos = tmp_pos + 1
            mid = mid + 1
        End If
    Loop
    
    Do While (left <= left_end)
        temp(tmp_pos) = numbers(left)
        left = left + 1
        tmp_pos = tmp_pos + 1
    Loop
    Do While (mid <= right)
        temp(tmp_pos) = numbers(mid)
        mid = mid + 1
        tmp_pos = tmp_pos + 1
    Loop
    
    For i = 0 To (num_elements - 1)
        numbers(right) = temp(right)
        right = right - 1
    Next i
End Sub

Public Sub QuickSort(numbers() As Integer, ByVal num As Integer)
    'do the quick sort
    q_sort numbers, 0, num - 1
End Sub

Public Sub q_sort(numbers() As Integer, ByVal left As Integer, ByVal right As Integer)
    Dim pivot As Integer
    Dim l_hold As Integer
    Dim r_hold As Integer
    
    'do the branching
    l_hold = left
    r_hold = right
    pivot = numbers(left)
    Do While (left < right)
        Do While ((numbers(right) >= pivot) And (left < right))
            right = right - 1
        Loop
        If (left <> right) Then
            numbers(left) = numbers(right)
            left = left + 1
        End If
        Do While ((numbers(left) <= pivot) And (left < right))
            left = left + 1
        Loop
        If (left <> right) Then
            numbers(right) = numbers(left)
            right = right - 1
        End If
    Loop
    numbers(left) = pivot
    pivot = left
    left = l_hold
    right = r_hold
    If (left < pivot) Then
        q_sort numbers, left, pivot - 1
    End If
    If (right > pivot) Then
        q_sort numbers, pivot + 1, right
    End If
End Sub

Public Sub RadixSort(numbers() As Integer, temp() As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim reps As Integer
    
    'find the number of repetitions needed
    reps = numbers(0)
    For i = 1 To (num - 1)
        If numbers(i) > reps Then
            reps = numbers(i)
        End If
    Next i
    reps = (Log(reps) / 2)
    
    'do the radix sort
    For i = 1 To reps
        If (i Mod 2) = 1 Then
            radix numbers(), temp(), num, i
        Else
            radix temp(), numbers(), num, i
        End If
    Next i
    
    'copy into proper array
    If (i Mod 2) = 0 Then
        For i = 0 To (num - 1)
            numbers(i) = temp(i)
        Next i
    End If
End Sub

Public Sub radix(numbers() As Integer, temp() As Integer, ByVal num As Integer, ByVal place As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim place2 As Integer
    Dim place3 As Integer
    
    'set the variables
    place2 = 10 ^ place
    place3 = place2 / 10
    k = 0
    
    'do the radix sort
    For j = 0 To 9
        For i = 0 To (num - 1)
            If ((numbers(i) Mod place2) \ place3) = j Then
                temp(k) = numbers(i)
                k = k + 1
            End If
        Next i
    Next j
End Sub

Public Sub OETSort(numbers() As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    
    'do the odd-even transposition sort
    For i = 0 To (num - 1)
        If ((i Mod 2) = 1) Then
            For j = 1 To (num - 1) Step 2
                If (j + 2 <= num) Then
                    If (numbers(j) > numbers(j + 1)) Then
                        temp = numbers(j)
                        numbers(j) = numbers(j + 1)
                        numbers(j + 1) = temp
                    End If
                End If
            Next j
        Else
            For j = 0 To (num - 1) Step 2
                If (j + 2 <= num) Then
                    If (numbers(j) > numbers(j + 1)) Then
                        temp = numbers(j)
                        numbers(j) = numbers(j + 1)
                        numbers(j + 1) = temp
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Public Sub Bitonic(numbers() As Integer, ByVal num As Integer)
    'start the sort
    sortup 0, num, numbers()
End Sub

Public Sub sortup(ByVal m As Integer, ByVal n As Integer, numbers() As Integer)
    'separate sort upwards
    If (n <> 1) Then
        sortup m, n \ 2, numbers()
        sortdown m + n \ 2, n \ 2, numbers()
        mergeup m, n \ 2, numbers()
    End If
End Sub

Public Sub sortdown(ByVal m As Integer, ByVal n As Integer, numbers() As Integer)
    'separate sort downwards
    If (n <> 1) Then
        sortup m, n \ 2, numbers()
        sortdown m + n \ 2, n \ 2, numbers()
        mergedown m, n \ 2, numbers()
    End If
End Sub

Public Sub mergeup(ByVal m As Integer, ByVal n As Integer, numbers() As Integer)
    Dim i As Integer
    Dim temp As Integer
    
    'merge sort upwards
    If (n <> 0) Then
        For i = 0 To n - 1
            If numbers(m + i) > numbers(m + i + n) Then
                temp = numbers(m + i)
                numbers(m + i) = numbers(m + i + n)
                numbers(m + i + n) = temp
            End If
        Next i
        mergeup m, n \ 2, numbers()
        mergeup m + n, n \ 2, numbers()
    End If
End Sub


Public Sub mergedown(ByVal m As Integer, ByVal n As Integer, numbers() As Integer)
    Dim i As Integer
    Dim temp As Integer
    
    'merge sort downwards
    If (n <> 0) Then
        For i = 0 To n - 1
            If numbers(m + i) < numbers(m + i + n) Then
                temp = numbers(m + i)
                numbers(m + i) = numbers(m + i + n)
                numbers(m + i + n) = temp
            End If
        Next i
        mergeup m, n \ 2, numbers()
        mergeup m + n, n \ 2, numbers()
    End If
End Sub

Public Sub BucketSort(numbers() As Integer, ByVal num As Integer, temp() As Integer)
    Dim i As Integer
    Dim j As Integer
    
    'clear temp
    For i = 0 To num
        temp(i) = 0
    Next i
    'count number of elements
    For i = 0 To (num - 1)
        temp(numbers(i)) = temp(numbers(i)) + 1
    Next i
    'reassign sorted values
    j = 0
    For i = 0 To num
        Do While (temp(i) > 0)
            numbers(j) = i
            j = j + 1
            temp(i) = temp(i) - 1
        Loop
    Next i
End Sub

Public Sub BinaryInsertionSort(numbers() As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim low As Integer
    Dim high As Integer
    Dim temp As Integer
    
    'do the binary insertion sort
    For i = 1 To (num - 1)
        temp = numbers(i)
        low = -1
        high = i
        Do While (high - low) > 1
            j = (high + low) \ 2
            If (temp < numbers(j)) Then
                high = j
            Else
                low = j
            End If
        Loop
        For j = i To (high + 1) Step -1
            numbers(j) = numbers(j - 1)
            numbers(high) = temp
        Next j
    Next i
End Sub

Public Sub BUMergeSort(list() As Integer, ByVal num As Integer, temp() As Integer)
    Dim l As Integer
    Dim r As Integer
    Dim m As Integer
    Dim n As Integer
    
    r = 0
    For n = 1 To (num - 1)
        For l = 0 To (num - 1) Step r + 1
            r = 1 + 2 * n - 1
            If (r > n) Then
                bumerge list(), l, l + n - 1, n, temp()
                If (l > 1) Then
                    bumerge list(), 1 - 2 * n, l - 1, n, temp()
                End If
            Else
                bumerge list(), l, l + n - 1, r, temp()
            End If
        Next l
    Next n
End Sub

Public Sub bumerge(list() As Integer, ByVal l As Integer, ByVal m As Integer, ByVal r As Integer, temp() As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    i = 1
    j = m + 1
    For k = 1 To r
        If (i > m) Then
            temp(k) = list(j)
            j = j + 1
        ElseIf (j > r) Then
            temp(k) = list(i)
            i = i + 1
        Else
            If list(i) < list(j) Then
                temp(k) = list(i)
                i = i + 1
            Else
                temp(k) = list(j)
                j = j + 1
            End If
        End If
    Next k
End Sub

Public Sub BubbleQuickSort(numbers() As Integer, ByVal num As Integer)
    'If there's more than 1 item, do the quick sort with bubble sort
    If (num <> 1) Then
        bubquick numbers(), 0, num - 1
    End If
End Sub

Public Sub bubquick(numbers() As Integer, ByVal left As Integer, ByVal right As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim v As Integer
    Dim temp As Integer
    Dim end_now As Boolean
    
    'if there's 5 or less items, do the bubble sort
    If (right - left <= 5) Then
        bub numbers(), right, left
    Else
        'otherwise, do the quick sort
        'using the tri-median method to find optimal placement for the pivot point
        i = (left + right) \ 2
        If (numbers(left) > numbers(i)) Then
            temp = numbers(left)
            numbers(left) = numbers(i)
            numbers(i) = temp
        End If
        If (numbers(left) > numbers(right)) Then
            temp = numbers(left)
            numbers(left) = numbers(right)
            numbers(right) = temp
        End If
        If (numbers(i) > numbers(right)) Then
            temp = numbers(i)
            numbers(i) = numbers(right)
            numbers(right) = temp
        End If
        j = right - 1
        temp = numbers(i)
        numbers(i) = numbers(j)
        numbers(j) = temp
        i = left
        v = numbers(j)
        end_now = False
        Do While Not end_now
            Do While (numbers(inc(i)) < v)
            Loop
            Do While (numbers(dec(j)) > v)
            Loop
            If (j < i) Then
                end_now = True
            Else
                temp = numbers(i)
                numbers(i) = numbers(j)
                numbers(j) = temp
            End If
        Loop
        temp = numbers(i)
        numbers(i) = numbers(right - 1)
        numbers(right - 1) = temp
        bubquick numbers(), left, j
        bubquick numbers(), i + 1, right
    End If
End Sub

Public Sub bub(numbers() As Integer, ByVal hi As Integer, ByVal lo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    Dim swapped As Boolean
    
    'special bubble sort
    For i = hi To lo Step -1
        swapped = False
        For j = lo To (i - 1)
            If (numbers(j) > numbers(j + 1)) Then
                temp = numbers(j)
                numbers(j) = numbers(j + 1)
                numbers(j + 1) = temp
                swapped = True
            End If
        Next j
        If Not swapped Then
            i = -5
        End If
    Next i
End Sub

Function inc(ByRef num As Integer) As Integer
    'increment a number before comparing
    num = num + 1
    inc = num
End Function

Function dec(ByRef num As Integer) As Integer
    'decrement a number before comparing
    num = num - 1
    dec = num
End Function

Public Sub GenerateRandom(list() As Integer, ByVal min As Integer, ByVal max As Integer, ByVal num As Integer)
    Dim i As Integer
    
    'randomize!!!!!
    Randomize ((min * max) Mod num)
    
    'fill array with random values
    For i = 0 To (num - 1)
        list(i) = (Rnd() * (max - min)) + min
    Next i
End Sub

Public Sub GenerateReverse(list() As Integer, ByVal min As Integer, ByVal max As Integer, ByVal num As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim temp(5000) As Integer
    
    'get random list
    GenerateRandom list(), min, max, num
    
    'sort the list using quick
    QuickSort list(), num
    
    'reverse the list
    j = num - 1
    For i = 0 To (num - 1)
        temp(i) = list(j)
        j = j - 1
    Next i
    For i = 0 To num - 1
        list(i) = temp(i)
    Next i
End Sub

Public Sub GenerateSorted(list() As Integer, ByVal min As Integer, ByVal max As Integer, ByVal num As Integer)
    'get random list
    GenerateRandom list(), min, max, num
    
    'sort the list using quick
    QuickSort list(), num
End Sub
