<?php

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Services\DashboardService;

class ApiDashboardController extends Controller
{
    protected $dashboardService;

    public function __construct(DashboardService $dashboardService)
    {
        $this->dashboardService = $dashboardService;
    }

    /**
     * List all available dashboards
     */
    public function index()
    {
        try {
            $dashboards = $this->dashboardService->listDashboards();
            
            // Add financial dashboards at the beginning of the list
            $financialDashboards = [
                [
                    'type' => 'financial',
                    'title' => 'Financial Dashboard',
                    'description' => 'Financial Card View with revenue breakdowns'
                ],
                [
                    'type' => 'financial-table',
                    'title' => 'Financial Table',
                    'description' => 'Financial Table View with monthly breakdowns'
                ]
            ];
            $dashboards = array_merge($financialDashboards, $dashboards);
            
            return response()->json([
                'success' => true,
                'data' => $dashboards,
                'count' => count($dashboards)
            ]);
        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'message' => 'Failed to load dashboards',
                'error' => $e->getMessage()
            ], 500);
        }
    }

    /**
     * Get a specific dashboard with all widgets
     */
    public function show(Request $request, string $type)
    {
        // Handle special financial dashboards
        if ($type === 'financial') {
            return $this->financial($request);
        }
        
        if ($type === 'financial-table') {
            return $this->financialTable($request);
        }
        
        $period = $request->input('period', 'all_time');
        $startDate = $request->input('start_date');
        $endDate = $request->input('end_date');
        
        // Get provider ID (default: 2087)
        $filters = [
            'franchise' => $request->input('franchise'),
        ];

        // Prioritize explicit dates over period
        if ($startDate && $endDate) {
            $filters['start_date'] = $startDate;
            $filters['end_date'] = $endDate;
        } elseif ($period !== 'all_time') {
            // Only use period if no explicit dates are provided
            $filters['period'] = $period;
        }

        try {
            $dashboardData = $this->dashboardService->getDashboardData($type, array_filter($filters));
            
            return response()->json([
                'success' => true,
                'data' => $dashboardData,
                'meta' => [
                    'type' => $type,
                    'filters' => $filters,
                    'period' => $period
                ]
            ]);
        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'message' => 'Failed to load dashboard',
                'error' => $e->getMessage(),
                'type' => $type
            ], 500);
        }
    }

    /**
     * Get dashboard data (alias for show)
     */
    public function getData(Request $request, string $type)
    {
        return $this->show($request, $type);
    }

    /**
     * Get financial dashboard data
     */
    public function financial(Request $request)
    {
        $clickhouse = app(\App\Services\ClickhouseService::class);
        $today = date('Y-m-d');
        $yesterday = date('Y-m-d', strtotime('-1 day'));
        $weekStart = date('Y-m-d', strtotime('monday this week'));
        $prevWeekStart = date('Y-m-d', strtotime('monday last week'));
        $prevWeekEnd = date('Y-m-d', strtotime('sunday last week'));
        
        // Month to date and last month
        $monthStart = date('Y-m-01');
        $lastMonthStart = date('Y-m-01', strtotime('first day of last month'));
        $lastMonthEnd = date('Y-m-t', strtotime('last day of last month'));
        
        // Year to date and last year
        $yearStart = date('Y-01-01');
        $lastYearStart = date('Y-01-01', strtotime('-1 year'));
        $lastYearEnd = date('Y-12-31', strtotime('-1 year'));

        // Saleable item types filter
        $saleableFilter = "(lowerUTF8(iid.item_type) IN ('product', 'service', 'class', 'membership', 'package', 'rental', 'giftcard', 'appointment', 'subscription') OR lowerUTF8(iid.item_type) LIKE 'misc%' OR lowerUTF8(iid.item_type) LIKE 'Misc%')";

        // Function to get revenue breakdown by type
        $getRevenueBreakdown = function($startDate, $endDate) use ($clickhouse, $saleableFilter) {
            $sql = "
 SELECT 
                    CASE 
                        WHEN lowerUTF8(iid.item_type) = 'membership' THEN 'membership'
                        WHEN lowerUTF8(iid.item_type) IN ('product', 'Product') THEN 'products'
                        WHEN lowerUTF8(iid.item_type) IN ('class', 'appointment', 'Appointment') THEN 'training'
                        WHEN lowerUTF8(iid.item_type) IN ('service', 'Service') THEN 'services'
                        WHEN lowerUTF8(iid.item_type) IN ('giftcard', 'GiftCard') THEN 'giftcards'
                        ELSE 'other'
                    END AS revenue_type,
                    SUM(iid.total_price) AS revenue
                FROM invoice_items_detail AS iid
                INNER JOIN invoice_details AS idt ON iid.invoice_id = idt.id
                WHERE idt.invoice_date BETWEEN toDate('{$startDate}') AND toDate('{$endDate}')
                    AND idt.status = 'active'
                    AND {$saleableFilter}
                GROUP BY revenue_type";

            try {
                $results = $clickhouse->select($sql);
                
                // Debug logging
                \Log::info('Financial dashboard query executed', [
                    'start_date' => $startDate,
                    'end_date' => $endDate,
                    'results_count' => count($results),
                    'results' => $results
                ]);
            } catch (\Exception $e) {
                \Log::error('Financial dashboard query failed', [
                    'sql' => $sql,
                    'error' => $e->getMessage(),
                    'start_date' => $startDate,
                    'end_date' => $endDate
                ]);
                return [
                    'membership' => 0,
                    'products' => 0,
                    'training' => 0,
                    'services' => 0,
                    'giftcards' => 0,
                    'total' => 0
                ];
            }
            
            $breakdown = [
                'membership' => 0,
                'products' => 0,
                'training' => 0,
                'services' => 0,
                'giftcards' => 0,
                'total' => 0
            ];

            if (empty($results)) {
                \Log::warning('Financial dashboard query returned empty results', [
                    'start_date' => $startDate,
                    'end_date' => $endDate,
                    'sql' => $sql
                ]);
            }

            foreach ($results as $row) {
                $type = $row['revenue_type'] ?? 'other';
                $revenue = (float)($row['revenue'] ?? 0);
                if (isset($breakdown[$type])) {
                    $breakdown[$type] = $revenue;
                }
                $breakdown['total'] += $revenue;
            }

            \Log::info('Financial dashboard breakdown calculated', [
                'start_date' => $startDate,
                'end_date' => $endDate,
                'breakdown' => $breakdown
            ]);
            
            // Debug: Check what membership item_types exist and test without saleableFilter
            if ($breakdown['membership'] == 0) {
                // First, check all item_types that might be memberships
                $debugSql1 = "
                    SELECT DISTINCT lowerUTF8(iid.item_type) as item_type, COUNT(*) as count, SUM(iid.total_price) as total_revenue
                    FROM invoice_items_detail AS iid
                    INNER JOIN invoice_details AS idt ON iid.invoice_id = idt.id
                    WHERE idt.invoice_date BETWEEN toDate('{$startDate}') AND toDate('{$endDate}')
                        AND idt.status = 'active'
                        AND (lowerUTF8(iid.item_type) LIKE '%membership%' OR lowerUTF8(iid.item_type) LIKE '%member%')
                    GROUP BY item_type
                ";
                
                // Also check what the saleableFilter is actually matching
                $debugSql2 = "
                    SELECT 
                        CASE 
                            WHEN lowerUTF8(iid.item_type) IN ('membership', 'memberships') THEN 'membership'
                            ELSE 'other'
                        END AS revenue_type,
                        lowerUTF8(iid.item_type) as actual_item_type,
                        COUNT(*) as count,
                        SUM(iid.total_price) as total_revenue
                    FROM invoice_items_detail AS iid
                    INNER JOIN invoice_details AS idt ON iid.invoice_id = idt.id
                    WHERE idt.invoice_date BETWEEN toDate('{$startDate}') AND toDate('{$endDate}')
                        AND idt.status = 'active'
                        AND {$saleableFilter}
                        AND (lowerUTF8(iid.item_type) LIKE '%membership%' OR lowerUTF8(iid.item_type) LIKE '%member%')
                    GROUP BY revenue_type, actual_item_type
                ";
                
                try {
                    $debugResults1 = $clickhouse->select($debugSql1);
                    $debugResults2 = $clickhouse->select($debugSql2);
                    \Log::warning('Membership revenue is zero - debugging', [
                        'start_date' => $startDate,
                        'end_date' => $endDate,
                        'all_membership_item_types' => $debugResults1,
                        'membership_in_saleable_filter' => $debugResults2,
                        'saleable_filter' => $saleableFilter
                    ]);
                } catch (\Exception $e) {
                    \Log::error('Debug query failed', ['error' => $e->getMessage()]);
                }
            }

            return $breakdown;
        };

        // Function to get monthly trend data (last 6 months)
        $getMonthlyTrend = function($endDate) use ($clickhouse, $saleableFilter) {
            $sixMonthsAgo = date('Y-m-01', strtotime($endDate . ' -5 months'));
            $sql = "
                 SELECT 
                    formatDateTime(toStartOfMonth(idt.invoice_date), '%b') AS month,
                    SUM(iid.total_price) AS revenue
                FROM invoice_items_detail AS iid
                INNER JOIN invoice_details AS idt ON iid.invoice_id = idt.id
                WHERE idt.invoice_date >= toDate('{$sixMonthsAgo}')
                    AND idt.invoice_date <= toDate('{$endDate}')
                    AND idt.status = 'active'
                    AND {$saleableFilter}
                GROUP BY toStartOfMonth(idt.invoice_date), month
                ORDER BY toStartOfMonth(idt.invoice_date) ASC
            ";

            $results = $clickhouse->select($sql);
            $trend = [
                'labels' => [],
                'data' => []
            ];

            foreach ($results as $row) {
                $trend['labels'][] = $row['month'] ?? '';
                $trend['data'][] = (float)($row['revenue'] ?? 0);
            }

            return $trend;
        };

        // Function to get location breakdown
        $getLocationBreakdown = function($startDate, $endDate) use ($clickhouse, $saleableFilter) {
            $sql = "
                SELECT 
                    idt.location,
                    SUM(iid.total_price) AS revenue
                FROM invoice_items_detail AS iid
                INNER JOIN invoice_details AS idt ON iid.invoice_id = idt.id
                WHERE idt.invoice_date BETWEEN toDate('{$startDate}') AND toDate('{$endDate}')
                    AND idt.status = 'active'
                    AND {$saleableFilter}
                GROUP BY idt.location
            ";

            $results = $clickhouse->select($sql);
            $locations = [
                'Warehouse' => 0,
                'Transport' => 0,
                'Retail' => 0
            ];

            foreach ($results as $row) {
                $location = $row['location'] ?? '';
                $revenue = (float)($row['revenue'] ?? 0);
                
                $locationLower = strtolower($location);
                if (strpos($locationLower, 'warehouse') !== false) {
                    $locations['Warehouse'] += $revenue;
                } elseif (strpos($locationLower, 'transport') !== false) {
                    $locations['Transport'] += $revenue;
                } elseif (strpos($locationLower, 'retail') !== false) {
                    $locations['Retail'] += $revenue;
                }
            }

            return $locations;
        };

        try {
            // Get data for all periods
            $todayData = $getRevenueBreakdown($today, $today);
            $todayData['trend'] = $getMonthlyTrend($today);
            $todayData['locations'] = $getLocationBreakdown($today, $today);
            
            $yesterdayData = $getRevenueBreakdown($yesterday, $yesterday);
            $yesterdayData['trend'] = $getMonthlyTrend($yesterday);
            $yesterdayData['locations'] = $getLocationBreakdown($yesterday, $yesterday);
            
            $weekToDateData = $getRevenueBreakdown($weekStart, $today);
            $weekToDateData['trend'] = $getMonthlyTrend($today);
            $weekToDateData['locations'] = $getLocationBreakdown($weekStart, $today);
            
            $prevWeekData = $getRevenueBreakdown($prevWeekStart, $prevWeekEnd);
            $prevWeekData['trend'] = $getMonthlyTrend($prevWeekEnd);
            $prevWeekData['locations'] = $getLocationBreakdown($prevWeekStart, $prevWeekEnd);
            
            $monthToDateData = $getRevenueBreakdown($monthStart, $today);
            $monthToDateData['trend'] = $getMonthlyTrend($today);
            $monthToDateData['locations'] = $getLocationBreakdown($monthStart, $today);
            
            $lastMonthData = $getRevenueBreakdown($lastMonthStart, $lastMonthEnd);
            $lastMonthData['trend'] = $getMonthlyTrend($lastMonthEnd);
            $lastMonthData['locations'] = $getLocationBreakdown($lastMonthStart, $lastMonthEnd);
            
            $yearToDateData = $getRevenueBreakdown($yearStart, $today);
            $yearToDateData['trend'] = $getMonthlyTrend($today);
            $yearToDateData['locations'] = $getLocationBreakdown($yearStart, $today);
            
            $lastYearData = $getRevenueBreakdown($lastYearStart, $lastYearEnd);
            $lastYearData['trend'] = $getMonthlyTrend($lastYearEnd);
            $lastYearData['locations'] = $getLocationBreakdown($lastYearStart, $lastYearEnd);

            return response()->json([
                'success' => true,
                'data' => [
                    'today' => $todayData,
                    'yesterday' => $yesterdayData,
                    'week_to_date' => $weekToDateData,
                    'prev_week' => $prevWeekData,
                    'month_to_date' => $monthToDateData,
                    'last_month' => $lastMonthData,
                    'year_to_date' => $yearToDateData,
                    'last_year' => $lastYearData
                ],
                'meta' => [
                    'date_ranges' => [
                        'today' => ['start' => $today, 'end' => $today],
                        'yesterday' => ['start' => $yesterday, 'end' => $yesterday],
                        'week_to_date' => ['start' => $weekStart, 'end' => $today],
                        'prev_week' => ['start' => $prevWeekStart, 'end' => $prevWeekEnd],
                        'month_to_date' => ['start' => $monthStart, 'end' => $today],
                        'last_month' => ['start' => $lastMonthStart, 'end' => $lastMonthEnd],
                        'year_to_date' => ['start' => $yearStart, 'end' => $today],
                        'last_year' => ['start' => $lastYearStart, 'end' => $lastYearEnd]
                    ]
                ]
            ]);
            
            // Test query to check if there's any data at all (for debugging)
            try {
                $testQuery = "SELECT COUNT(*) as total_invoices FROM invoice_details WHERE status = 'active' LIMIT 1";
                $testResult = $clickhouse->select($testQuery);
                \Log::info('Financial dashboard test query', ['result' => $testResult]);
            } catch (\Exception $e) {
                \Log::error('Financial dashboard test query failed', ['error' => $e->getMessage()]);
            }
        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'message' => 'Failed to load financial dashboard data',
                'error' => $e->getMessage()
            ], 500);
        }
    }

    /**
     * Get financial table view data
     */
    public function financialTable(Request $request)
    {
        $clickhouse = app(\App\Services\ClickhouseService::class);
        // Get last 12 months instead of current year
        $endDate = date('Y-m-d');
        $startDate = date('Y-m-01', strtotime('-11 months')); // Start from 12 months ago, beginning of that month

        // Saleable item types filter
        $saleableFilter = "(lowerUTF8(iid.item_type) IN ('product', 'service', 'class', 'membership', 'package', 'rental', 'giftcard', 'appointment', 'subscription') OR lowerUTF8(iid.item_type) LIKE 'misc%' OR lowerUTF8(iid.item_type) LIKE 'Misc%')";

        // Query monthly revenue breakdown by type
        $sql = "
            SELECT 
                toStartOfMonth(idt.invoice_date) AS month_start,
                formatDateTime(toStartOfMonth(idt.invoice_date), '%b') AS month_abbr,
                formatDateTime(toStartOfMonth(idt.invoice_date), '%Y') AS year,
                formatDateTime(toStartOfMonth(idt.invoice_date), '%b %Y') AS month_year,
                CASE 
                    WHEN lowerUTF8(iid.item_type) = 'membership' THEN 'memberships'
                    WHEN lowerUTF8(iid.item_type) IN ('product', 'Product') THEN 'products'
                    WHEN lowerUTF8(iid.item_type) IN ('service', 'Service') THEN 'services'
                    WHEN lowerUTF8(iid.item_type) IN ('class', 'appointment', 'Appointment') THEN 'training'
                    WHEN lowerUTF8(iid.item_type) = 'package' THEN 'packages'
                    ELSE 'other'
                END AS revenue_type,
                SUM(iid.total_price) AS revenue
            FROM invoice_items_detail AS iid
            INNER JOIN invoice_details AS idt ON iid.invoice_id = idt.id
            WHERE idt.invoice_date >= toDate('{$startDate}') AND idt.invoice_date <= toDate('{$endDate}')
                AND idt.status = 'active'
                AND {$saleableFilter}
            GROUP BY month_start, month_abbr, year, month_year, revenue_type
            ORDER BY month_start ASC
        ";

        try {
            $results = $clickhouse->select($sql);
            
            // Debug: Log raw results
            \Log::info('Financial table query results', [
                'start_date' => $startDate,
                'end_date' => $endDate,
                'results_count' => count($results),
                'sample_results' => array_slice($results, 0, 10),
                'membership_results' => array_filter($results, function($row) {
                    return ($row['revenue_type'] ?? '') === 'memberships';
                })
            ]);

            // Initialize monthly data structure - use month_year as key
            $monthlyData = [];
            
            // Populate monthly data using month_year as key
            foreach ($results as $row) {
                $monthYear = $row['month_year'] ?? '';
                $type = $row['revenue_type'] ?? 'other';
                $revenue = (float)($row['revenue'] ?? 0);
                
                // Normalize revenue_type (handle cases where provider ID suffix might be appended)
                // The query sometimes returns "memberships_2087" instead of "memberships"
                $normalizedType = $type;
                if (strpos($type, 'membership') === 0 || strpos($type, 'Membership') === 0) {
                    $normalizedType = 'memberships';
                }
                
                // Skip if no month_year
                if (empty($monthYear)) {
                    continue;
                }
                
                // Initialize month if not exists
                if (!isset($monthlyData[$monthYear])) {
                    $monthlyData[$monthYear] = [
                        'month_abbr' => $row['month_abbr'] ?? '',
                        'year' => $row['year'] ?? '',
                        'month_year' => $monthYear,
                        'month_start' => $row['month_start'] ?? '', // Store for sorting
                        'memberships' => 0,
                        'products' => 0,
                        'services' => 0,
                        'training' => 0,
                        'packages' => 0,
                        'total' => 0
                    ];
                }
                
                // Add revenue to the appropriate type
                if (isset($monthlyData[$monthYear][$normalizedType])) {
                    // Assign revenue directly (GROUP BY ensures one row per month+type combination)
                    $monthlyData[$monthYear][$normalizedType] = $revenue;
                    // Recalculate total for this month after each assignment
                    $monthlyData[$monthYear]['total'] = 
                        ($monthlyData[$monthYear]['memberships'] ?? 0) +
                        ($monthlyData[$monthYear]['products'] ?? 0) +
                        ($monthlyData[$monthYear]['services'] ?? 0) +
                        ($monthlyData[$monthYear]['training'] ?? 0) +
                        ($monthlyData[$monthYear]['packages'] ?? 0);
                } else {
                    // Log if type doesn't match expected keys
                    \Log::warning('Financial table: Unexpected revenue_type', [
                        'month_year' => $monthYear,
                        'original_type' => $type,
                        'normalized_type' => $normalizedType,
                        'revenue' => $revenue,
                        'available_keys' => array_keys($monthlyData[$monthYear])
                    ]);
                }
            }
            
            // Sort by month_start date (stored in month_year format like "Jan 2024")
            uksort($monthlyData, function($a, $b) use ($monthlyData) {
                $dateA = isset($monthlyData[$a]['month_start']) ? strtotime($monthlyData[$a]['month_start']) : strtotime($a);
                $dateB = isset($monthlyData[$b]['month_start']) ? strtotime($monthlyData[$b]['month_start']) : strtotime($b);
                return $dateA <=> $dateB;
            });
            
            // Debug logging for membership
            $totalMemberships = 0;
            foreach ($monthlyData as $month => $data) {
                $totalMemberships += $data['memberships'] ?? 0;
            }
            
            \Log::info('Financial table: Monthly data summary', [
                'total_months' => count($monthlyData),
                'total_memberships_revenue' => $totalMemberships,
                'monthly_data' => $monthlyData
            ]);
            
            if ($totalMemberships == 0) {
                // Debug query to see what membership item_types exist
                $debugSql = "
                    SELECT 
                        lowerUTF8(iid.item_type) as item_type,
                        COUNT(*) as count,
                        SUM(iid.total_price) as total_revenue,
                        MIN(idt.invoice_date) as first_date,
                        MAX(idt.invoice_date) as last_date
                    FROM invoice_items_detail AS iid
                    INNER JOIN invoice_details AS idt ON iid.invoice_id = idt.id
                    WHERE idt.invoice_date >= toDate('{$startDate}') AND idt.invoice_date <= toDate('{$endDate}')
                        AND idt.status = 'active'
                        AND (lowerUTF8(iid.item_type) LIKE '%membership%' OR lowerUTF8(iid.item_type) LIKE '%member%')
                    GROUP BY item_type
                ";
                try {
                    $debugResults = $clickhouse->select($debugSql);
                    \Log::warning('Financial table: Membership revenue is zero - debugging', [
                        'start_date' => $startDate,
                        'end_date' => $endDate,
                        'found_item_types' => $debugResults,
                        'monthly_data_sample' => array_slice($monthlyData, 0, 3, true),
                        'all_revenue_types_in_results' => array_unique(array_column($results, 'revenue_type'))
                    ]);
                } catch (\Exception $e) {
                    \Log::error('Financial table debug query failed', ['error' => $e->getMessage()]);
                }
            }
            
            // Debug logging
            \Log::info('Financial table query executed', [
                'start_date' => $startDate,
                'end_date' => $endDate,
                'results_count' => count($results),
                'results' => $results,
                'monthly_data' => $monthlyData
            ]);

            // Calculate column totals
            $columnTotals = [
                'memberships' => 0,
                'products' => 0,
                'services' => 0,
                'training' => 0,
                'packages' => 0,
                'total' => 0
            ];

            foreach ($monthlyData as $month => $data) {
                foreach ($columnTotals as $key => $value) {
                    if ($key !== 'total') {
                        $columnTotals[$key] += $data[$key] ?? 0;
                    }
                }
            }
            $columnTotals['total'] = array_sum(array_slice($columnTotals, 0, -1));

            return response()->json([
                'success' => true,
                'data' => [
                    'monthly_data' => $monthlyData,
                    'column_totals' => $columnTotals
                ],
                'meta' => [
                    'period' => 'last_12_months',
                    'start_date' => $startDate,
                    'end_date' => $endDate
                ]
            ]);
        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'message' => 'Failed to load financial table data',
                'error' => $e->getMessage()
            ], 500);
        }
    }
}

