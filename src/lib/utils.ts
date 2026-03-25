import { type ClassValue, clsx } from "clsx"
import { twMerge } from "tailwind-merge"

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}

export function formatDate(date: string | Date): string {
  return new Date(date).toLocaleDateString("en-US", {
    year: "numeric",
    month: "short",
    day: "numeric",
  })
}

export function formatDateTime(date: string | Date): string {
  return new Date(date).toLocaleDateString("en-US", {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  })
}

export function getDaysUntilExpiry(expiryDate: string | Date): number {
  const now = new Date();
  const expiry = new Date(expiryDate);
  const timeDiff = expiry.getTime() - now.getTime();
  return Math.ceil(timeDiff / (1000 * 3600 * 24));
}

export function getExpirationStatus(expiryDate: string | Date): string {
  const daysLeft = getDaysUntilExpiry(expiryDate);

  if (daysLeft <= 0) {
    return "expired";
  } else if (daysLeft <= 5) {
    return "expiring-soon";
  } else if (daysLeft <= 15) {
    return "expiring-warning";
  } else {
    return "valid";
  }
}

export function getFormStatus(form: any, latestSubmission: any = null): {
  status: 'pending' | 'active' | 'warning' | 'expired';
  statusText: string;
  daysLeft: number;
  colorClass: string;
  bgColorClass: string;
  borderColorClass: string;
} {
  // If no submission exists, form is pending (needs first submission)
  if (!latestSubmission) {
    return {
      status: 'pending',
      statusText: 'Pending Submission',
      daysLeft: 0,
      colorClass: 'text-red-600',
      bgColorClass: 'bg-red-50',
      borderColorClass: 'border-red-200'
    };
  }

  // Calculate expiry date from submission date + validity period
  const submissionDate = new Date(latestSubmission.submittedAt || latestSubmission.createdAt);
  const expiryDate = new Date(submissionDate);
  expiryDate.setDate(expiryDate.getDate() + (form.validityPeriod || 30));
  
  const daysLeft = getDaysUntilExpiry(expiryDate);

  // Form has expired (needs resubmission)
  if (daysLeft <= 0) {
    return {
      status: 'expired',
      statusText: 'Expired - Resubmit Required',
      daysLeft: Math.abs(daysLeft),
      colorClass: 'text-red-600',
      bgColorClass: 'bg-red-50',
      borderColorClass: 'border-red-200'
    };
  }
  
  // Form expires today or tomorrow (warning)
  if (daysLeft <= 2) {
    return {
      status: 'warning',
      statusText: daysLeft === 1 ? 'Due Tomorrow' : `${daysLeft} days left`,
      daysLeft,
      colorClass: 'text-yellow-600',
      bgColorClass: 'bg-yellow-50',
      borderColorClass: 'border-yellow-200'
    };
  }
  
  // Form is still valid (active)
  return {
    status: 'active',
    statusText: `${daysLeft} days left`,
    daysLeft,
    colorClass: 'text-green-600',
    bgColorClass: 'bg-green-50',
    borderColorClass: 'border-green-200'
  };
}

export function getDueDateStatus(dueDate: string | Date): {
  status: 'due-today' | 'due-soon' | 'due-later';
  colorClass: string;
  bgColorClass: string;
  borderColorClass: string;
} {
  const daysLeft = getDaysUntilExpiry(dueDate);

  if (daysLeft <= 0) {
    return {
      status: 'due-today',
      colorClass: 'text-red-600',
      bgColorClass: 'bg-red-50',
      borderColorClass: 'border-red-200'
    };
  } else if (daysLeft <= 2) {
    return {
      status: 'due-soon',
      colorClass: 'text-yellow-600',
      bgColorClass: 'bg-yellow-50',
      borderColorClass: 'border-yellow-200'
    };
  } else {
    return {
      status: 'due-later',
      colorClass: 'text-green-600',
      bgColorClass: 'bg-green-50',
      borderColorClass: 'border-green-200'
    };
  }
}