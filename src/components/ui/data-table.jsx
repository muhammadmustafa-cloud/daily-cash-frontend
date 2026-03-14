import React, { useState, useMemo } from "react"
import { DataGrid } from "@mui/x-data-grid"
import { Box } from "@mui/material"
import { Edit, Trash2, Eye } from "lucide-react"
import { IconButton } from "@mui/material"

export function DataTable({
  columns,
  data,
  onEdit,
  onView,
  onDelete,
  loading = false,
  pageSize = 10,
  ...props
}) {
  // Transform columns to MUI format
  const gridColumns = columns.map((col) => ({
    field: col.key,
    headerName: col.label,
    flex: 1,
    minWidth: 150,
    renderCell: col.render
      ? (params) => col.render(params.value, params.row)
      : undefined,
  }))

  // Add actions column if onEdit, onView, or onDelete provided
  if (onEdit || onView || onDelete) {
    gridColumns.push({
      field: "actions",
      headerName: "Actions",
      width: onView && onDelete ? 140 : 120,
      sortable: false,
      filterable: false,
      renderCell: (params) => (
        <Box sx={{ display: "flex", gap: 1 }}>
          {onView && (
            <IconButton
              size="small"
              onClick={() => onView(params.row)}
              color="primary"
              title="View"
            >
              <Eye className="h-4 w-4" />
            </IconButton>
          )}
          {onEdit && (
            <IconButton
              size="small"
              onClick={() => onEdit(params.row)}
              color="primary"
              title="Edit"
            >
              <Edit className="h-4 w-4" />
            </IconButton>
          )}
          {onDelete && (
            <IconButton
              size="small"
              onClick={() => onDelete(params.row)}
              color="error"
              title="Delete"
            >
              <Trash2 className="h-4 w-4" />
            </IconButton>
          )}
        </Box>
      ),
    })
  }

  // Transform data to include id field
  const rows = data.map((row, index) => ({
    id: row._id || row.id || index,
    ...row,
  }))

  return (
    <Box sx={{ height: 600, width: "100%", '& .MuiDataGrid-root': { border: 'none' } }}>
      <DataGrid
        rows={rows}
        columns={gridColumns}
        loading={loading}
        pageSize={pageSize}
        rowsPerPageOptions={[10, 25, 50, 100]}
        disableSelectionOnClick
        sx={{
          border: "none",
          "& .MuiDataGrid-cell": {
            borderBottom: "1px solid #e5e7eb",
          },
          "& .MuiDataGrid-columnHeaders": {
            backgroundColor: "#f9fafb",
            borderBottom: "2px solid #e5e7eb",
            fontWeight: 600,
          },
          "& .MuiDataGrid-footerContainer": {
            borderTop: "1px solid #e5e7eb",
          },
          "& .MuiDataGrid-row:hover": {
            backgroundColor: "#f9fafb",
          },
        }}
        {...props}
      />
    </Box>
  )
}
