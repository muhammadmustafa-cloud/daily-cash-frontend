import { useState, useEffect } from 'react';
import { Trash2, Eye, ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight } from 'lucide-react';
import { Button } from './ui/Button';

export const EntriesTable = ({
  entries = [],
  entryType,
  onEdit,
  onDelete,
  onViewImage,
  onPageChange,
  page = 1,
  total = 0,
  limit = 20,
  showPagination = true,
  emptyMessage = 'No entries yet.'
}) => {
  const totalPages = Math.ceil(total / limit);
  const startItem = (page - 1) * limit + 1;
  const endItem = Math.min(page * limit, total);

  const handlePageChange = (newPage) => {
    if (newPage >= 1 && newPage <= totalPages) {
      onPageChange(newPage);
    }
  };

  if (entries.length === 0) {
    return (
      <div className="text-center py-8 text-gray-500">
        {emptyMessage}
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="overflow-x-auto">
        <table className="w-full">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-4 py-3 text-left text-xs font-semibold text-gray-700 uppercase">Name</th>
              <th className="px-4 py-3 text-left text-xs font-semibold text-gray-700 uppercase">Description</th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-700 uppercase">Amount</th>
              <th className="px-4 py-3 text-center text-xs font-semibold text-gray-700 uppercase">Image</th>
              <th className="px-4 py-3 text-center text-xs font-semibold text-gray-700 uppercase">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-200">
            {entries.map((entry, index) => (
              <tr key={entry._id || index} className="hover:bg-gray-50">
                <td className="px-4 py-3 text-sm text-gray-900">{entry.name}</td>
                <td className="px-4 py-3 text-sm text-gray-600">{entry.description || '-'}</td>
                <td className="px-4 py-3 text-sm text-right font-semibold text-gray-900">
                  {new Intl.NumberFormat('en-PK', {
                    style: 'currency',
                    currency: 'PKR',
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                  }).format(entry.amount)}
                </td>
                <td className="px-4 py-3 text-center">
                  {entry.image ? (
                    <button
                      onClick={() => onViewImage(entry.image)}
                      className="text-blue-600 hover:text-blue-800 transition-colors"
                      title="View image"
                      aria-label={`View image for ${entry.name}`}
                    >
                      <Eye className="h-5 w-5" />
                    </button>
                  ) : (
                    <span className="text-gray-400">-</span>
                  )}
                </td>
                <td className="px-4 py-3 text-center">
                  <div className="flex justify-center gap-2">
                    <button
                      onClick={() => onEdit(entry, entryType)}
                      className="text-blue-600 hover:text-blue-800 text-sm"
                      aria-label={`Edit ${entry.name}`}
                    >
                      Edit
                    </button>
                    <button
                      onClick={() => onDelete(entry, entryType)}
                      className="text-red-600 hover:text-red-800 text-sm"
                      aria-label={`Delete ${entry.name}`}
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {showPagination && totalPages > 1 && (
        <div className="flex items-center justify-between px-2 py-3">
          <div className="text-sm text-gray-700">
            Showing <span className="font-medium">{startItem}</span> to{' '}
            <span className="font-medium">{endItem}</span> of{' '}
            <span className="font-medium">{total}</span> entries
          </div>
          <div className="flex items-center space-x-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => handlePageChange(1)}
              disabled={page === 1}
              aria-label="First page"
            >
              <ChevronsLeft className="h-4 w-4" />
            </Button>
            <Button
              variant="outline"
              size="sm"
              onClick={() => handlePageChange(page - 1)}
              disabled={page === 1}
              aria-label="Previous page"
            >
              <ChevronLeft className="h-4 w-4" />
            </Button>
            <span className="px-3 py-1 text-sm text-gray-700">
              Page {page} of {totalPages}
            </span>
            <Button
              variant="outline"
              size="sm"
              onClick={() => handlePageChange(page + 1)}
              disabled={page >= totalPages}
              aria-label="Next page"
            >
              <ChevronRight className="h-4 w-4" />
            </Button>
            <Button
              variant="outline"
              size="sm"
              onClick={() => handlePageChange(totalPages)}
              disabled={page >= totalPages}
              aria-label="Last page"
            >
              <ChevronsRight className="h-4 w-4" />
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};

export default EntriesTable;
