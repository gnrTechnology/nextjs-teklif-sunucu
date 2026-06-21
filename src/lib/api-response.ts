import { NextResponse } from "next/server";

export function jsonResponse(data: unknown, status = 200): NextResponse {
  return NextResponse.json(data, { status });
}

export function errorResponse(
  message: string,
  status = 400,
  extra: Record<string, unknown> = {},
): NextResponse {
  return NextResponse.json({ success: false, message, ...extra }, { status });
}
