# 工廠排程 - 使用NSGA-II

![image](https://github.com/clay-cchuang/NSGA-II/blob/main/gantt.png)
## 簡介

這個專案是一個用於工廠排程管理的系統，與之前[多目標基因演算法](https://github.com/yocolayo/Genetic-Algorithm)的限制相同，本次將專注於使用NSGA-II（Non-dominated Sorting Genetic Algorithm II）來解決這個問題。

NSGA-II是一種多目標優化算法，用於解決具有多個目標函數的優化問題。在這個排程管理專案中，NSGA-II的目標是找到一組工單排程，使得以下目標最佳化：

1. **最大化生產線使用時間 (Maxspan)**：確保生產線被充分利用，最大限度地減少閒置時間。

2. **最小化設置時間 (Setup Time)**：儘量減少在不同工單之間切換所需的設置時間，從而提高生產效率。

3. **最小化工單的遲繳時間 (Tardiness)**：確保所有工單都在其交期之前完成，減少生產延遲。

NSGA-II運作原理：

1. **初始化**：首先，NSGA-II創建一個初始的工單排程集合，這些排程是基於隨機生成的工單順序。

2. **計算適應度值**：針對每個排程，NSGA-II計算每個工單的相關目標函數，包括Maxspan、Setup Time和Tardiness。這些目標函數用於評估每個排程的品質。

3. **非支配排序**：NSGA-II使用非支配排序方法將排程分為不同的前沿（fronts）。前沿包含了不同等級的排程，根據工單排程的優劣進行排序。優秀的排程被歸為第一前沿，次優的排程被歸為第二前沿，以此類推。

4. **計算擁擠度**：在同一前沿中，NSGA-II使用擁擠度來評估排程的多樣性。擁擠度是根據工單排程在目標空間中的分佈來計算的，目的是確保生成多樣的優化解。

5. **選擇和交配**：NSGA-II選擇優秀的排程，並使用交配操作創建新的排程。交配操作通常包括片段的互換和隨機變異。

6. **迭代優化**：上述步驟進行多代迭代，不斷優化工單排程，以獲得最佳的優化解。

最終，NSGA-II產生了一組優化的工單排程，這些排程考慮了Maxspan、Setup Time和Tardiness等多個目標，從而幫助工廠實現更高效的生產流程管理。此方法的優勢在於能夠處理多目標排程問題，並找到一組優秀的解決方案，以平衡不同的生產需求。


## 設定

1. 在專案根目錄中，可以直接下載 `WIP.xlsx` 這個 Excel 檔案，也可更動檔案中的數值。

2. 在程式碼中，可以定義不同機台和產品類型的加工時間。以下是本專案加工時間的定義：

   | 產品類型 | 機台 | 填充時間 (秒) | 貼標時間 (秒) |
   |----------|------|----------------|----------------|
   | A        | 330  | 7.2            | 0              |
   | A        | 310  | 18.0           | 0              |
   | A        | 380  | 12.0           | 0              |
   | A        | 500  | 7.2            | 0              |
   | A        | 600  | 7.2            | 0              |
   | B        | 6000 | 20.0           | 0              |
   | B        | 20000| 25.0           | 0              |
   | B        | 17250| 10.7           | 0              |
   | C        | 240  | 21.4           | 0              |
   | D        | 330  | 7.2            | 90.0           |
   | D        | 380  | 12.0           | 90.0           |
   | D        | 500  | 7.2            | 14.4           |


## 執行

要運行這個工廠排程管理系統，只需運行主程式。系統將使用預設的排程演算法 (EDD) 來排程工單。

```bash
python main.py
```

## 結果

排程結果將以 Gantt 圖的形式呈現，顯示了每個工單的開始和完成時間，以及它們在不同機台上的排程情況，同時也能將排序後的工單印出，確保排序正確。
